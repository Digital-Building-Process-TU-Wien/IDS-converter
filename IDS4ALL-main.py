import argparse
import os
from datetime import date
import pandas as pd
from ifctester import ids
from custom_functions import *

# Default settings for Excel to IDS direction
excel_path_default = "./Excel-files/"
excel_name_default = "Requirements-openBIM-building-permit-City-of-Vienna-extended"
output_path_default = "./output/"

def detect_direction(fmt):
    """Detect conversion direction from file format.
    
    :param fmt: File format/extension (e.g., '.xlsx', '.ids')
    :type fmt: str
    :return: Direction ('excel_to_ids' or 'ids_to_excel')
    :rtype: str
    :raises ValueError: If format is unknown
    """
    fmt_lower = fmt.lower()
    if fmt_lower in ('.xlsx', '.xls'):
        return 'excel_to_ids'
    elif fmt_lower == '.ids':
        return 'ids_to_excel'
    else:
        raise ValueError(f"Unknown file format: {fmt}. Expected .xlsx, .xls, or .ids")

def first_set(*values, default=None):
    """Return the first non-None value from the provided arguments.
    
    Useful for implementing fallback chains: try new argument, fall back to legacy,
    then use default if nothing is set.
    
    :param values: Arguments to check in order
    :param default: Default value if all arguments are None
    :return: First non-None value or default
    """
    return next((v for v in values if v is not None), default)

def parse_arguments():
    """Parse command-line arguments supporting both directions."""
    parser = argparse.ArgumentParser(description="Convert between Excel and IDS formats.")
    
    # Generic arguments
    parser.add_argument(
        "--input_path",
        default=None,
        help="Path to the directory containing the input file.",
    )
    parser.add_argument(
        "--input_name",
        default=None,
        help="Name of the input file (without extension).",
    )
    parser.add_argument(
        "--input_format",
        default=None,
        help="Format of the input file (.xlsx, .xls, or .ids). Auto-detected from extension if not provided.",
    )
    parser.add_argument(
        "--output_path",
        default=output_path_default,
        help="Path to the output directory for generated files.",
    )
    
    # Legacy aliases for backwards compatibility
    parser.add_argument(
        "--excel_path",
        default=None,
        help="(Deprecated: use --input_path) Path to the directory containing the Excel file.",
    )
    parser.add_argument(
        "--excel_name",
        default=None,
        help="(Deprecated: use --input_name) Name of the Excel file (without extension).",
    )
    parser.add_argument(
        "--excel_format",
        default=None,
        help="(Deprecated: use --input_format) Format of the Excel file.",
    )
    parser.add_argument(
        "--ids_path",
        default=None,
        help="(Deprecated: use --input_path) Path to the directory containing the IDS file.",
    )
    parser.add_argument(
        "--ids_name",
        default=None,
        help="(Deprecated: use --input_name) Name of the IDS file (without extension).",
    )
    parser.add_argument(
        "--ids_format",
        default=None,
        help="(Deprecated: use --input_format) Format of the IDS file.",
    )

    args = parser.parse_args()
    
    # Check if using defaults before filling them in
    using_defaults = args == parser.parse_args([])
    
    # Handle legacy arguments and set defaults
    args.input_name = first_set(args.input_name, args.excel_name, args.ids_name,
                                default=excel_name_default)
    args.input_path = first_set(args.input_path, args.excel_path, args.ids_path,
                                default=excel_path_default)
    args.input_format = first_set(args.input_format, args.excel_format, args.ids_format,
                                  default=".xlsx")
    
    if using_defaults:
        print("Default values were used, check -h for override parameters")

    return args

def get_metadata(excel_path, excel_name, excel_format, sheet_name):
    """Extract metadata from the Excel file."""
    full_excel_path = os.path.join(excel_path, excel_name + excel_format)
    
    # Load Excel data
    data = pd.read_excel(full_excel_path, sheet_name='IDS4ALL', usecols=[0, 1], header=None, skiprows=1)
    
    # Convert data to a dictionary
    data_dict = {key: value for key, value in zip(data[0], data[1]) if pd.notna(value)}
    
    # Validate required metadata
    required_metadata = ['Sheet name', 'IFC version']
    for meta in required_metadata:
        if meta not in data_dict:
            raise Exception(f'{meta} is not defined in the IDS4ALL sheet of the Excel')
    
    sheet_name = data_dict['Sheet name']
    ifc_version = data_dict['IFC version'].replace(' ','').replace(',','|')
    separate_by = data_dict['File separators'].replace(' ','').split(',') if 'File separators' in data_dict else []
    skipped_rows = data_dict['Skipped rows'] if 'Skipped rows' in data_dict else 0
    is_entity_based_app = True if 'Entity-based applicability' in data_dict and data_dict['Entity-based applicability'].lower() == 'yes' else False
    
    return sheet_name, ifc_version, separate_by, skipped_rows, is_entity_based_app, data_dict

def process_excel_data(excel_path, excel_name, excel_format, sheet_name, separate_by, skipped_rows, ifc_version, is_entity_based_app):
    """Process Excel data and convert to specifications."""
    full_excel_path = os.path.join(excel_path, excel_name + excel_format)
    
    # Convert Excel data to specifications
    excel_data = excel_to_spec_list(full_excel_path, sheet_name, separate_by, skipped_rows, ifc_version, is_entity_based_app)
    
    # Separate specifications by general data
    separated_excel_data = separate_specs_by_generaldata(excel_data)
    
    return separated_excel_data

def create_ids_files(separated_excel_data, data_dict, output_path, excel_name, sheet_name):
    """Create IDS files from specifications."""
    for key in separated_excel_data.keys():
        sep_data = separated_excel_data[key]
        
        # Construct purpose and milestone strings
        string_milestone = ''
        string_purpose = ''
        if len(sep_data['general']['Phase']) != 0: string_milestone += ', '.join(sep_data['general']['Phase'])
        if len(sep_data['general']['Role']) != 0: string_purpose += 'Role: ' + ', '.join(sep_data['general']['Role']) + '; '
        if len(sep_data['general']['Usecase']) != 0: string_purpose += 'Usecase: ' + ', '.join(sep_data['general']['Usecase']) + '; '
        string_purpose = string_purpose[0:len(string_purpose)-2]
        
        # Create IDS object
        my_ids = ids.Ids(
            title=data_dict['Title'] if 'Title' in data_dict else 'Not Defined',
            copyright=data_dict['Copyright'] if 'Copyright' in data_dict else None,
            version=data_dict['Version'] if 'Version' in data_dict else None,
            description=data_dict['Description'] if 'Description' in data_dict else None,
            author=data_dict['Author'] if 'Author' in data_dict else None,
            date=date.today().strftime('%Y-%m-%d'),
            purpose=string_purpose if string_purpose != '' else None,
            milestone=string_milestone if string_milestone != '' else None,
        )
        
        # Create IDS specifications
        create_ids_specifications(my_ids, sep_data['specs'])
        
        try:
            # Prepare output path and filename
            key = key.replace('/','-')
            os.makedirs(output_path, exist_ok=True)
            filename = excel_name + '_' + sheet_name + key + '.ids'
            output_path_full = os.path.join(output_path, filename)
            
            # Save IDS file
            result = my_ids.to_xml(output_path_full)
            
            # Add comment to XML
            add_comment_to_xml(output_path_full, ' Created with the IDS4ALL Converter developed by Simon Fischer, Harald Urban, Konstantin Höbart, and Christian Schranz of TU Wien Research Unit Digital Building Process (https://www.tuwien.at/en/cee/ibb/zdb). ')
            print('XML created')
            print(f'Output file: {output_path_full}')
        except Exception as error:
            print('Error: ', error)

def run_excel_to_ids(excel_path, excel_name, excel_format, output_path):
    """Convert Excel file to IDS file(s).
    
    :param excel_path: Path to directory containing Excel file
    :param excel_name: Name of Excel file without extension
    :param excel_format: Format of Excel file (.xlsx or .xls)
    :param output_path: Path to output directory
    """
    print(f"Excel file path: {os.path.join(excel_path, excel_name + excel_format)}") 
    print(f"Output path: {output_path}")
    print()
    
    # Extract metadata
    sheet_name, ifc_version, separate_by, skipped_rows, is_entity_based_app, data_dict = get_metadata(excel_path, excel_name, excel_format, 'IDS4ALL')
    
    # Process Excel data
    separated_excel_data = process_excel_data(excel_path, excel_name, excel_format, sheet_name, separate_by, skipped_rows, ifc_version, is_entity_based_app)
    
    # Create IDS files
    create_ids_files(separated_excel_data, data_dict, output_path, excel_name, sheet_name)

def run_ids_to_excel(ids_path, ids_name, ids_format, output_path):
    """Convert IDS file to Excel file.
    
    :param ids_path: Path to directory containing IDS file
    :param ids_name: Name of IDS file without extension
    :param ids_format: Format of IDS file (.ids)
    :param output_path: Path to output directory
    """
    print(f"IDS file path: {os.path.join(ids_path, ids_name + ids_format)}") 
    print(f"Output path: {output_path}")
    print()
    
    # Create Excel file from IDS
    convert_IDS_to_Excel(ids_path, ids_name, ids_format, output_path)

def main():
    args = parse_arguments()
    
    input_path = args.input_path
    input_name = args.input_name
    input_format = args.input_format
    output_path = args.output_path
    
    # Detect direction from format
    direction = detect_direction(input_format)
    
    if direction == 'excel_to_ids':
        run_excel_to_ids(input_path, input_name, input_format, output_path)
    else:  # ids_to_excel
        run_ids_to_excel(input_path, input_name, input_format, output_path)

if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Error: {error}")
        raise SystemExit(1)