import argparse
import os
from datetime import date
import pandas as pd
from ifctester import ids
from custom_functions import *

#Default settings
excel_path_default = "./Excel-files/"
excel_name_default = "Requirements-openBIM-building-permit-City-of-Vienna-extended"
output_path_default = "./output/"

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description="Generate IDS files from Excel data.")
    
    parser.add_argument(
        "--excel_path",
        default=excel_path_default,
        help="Path to the directory containing the Excel file.",
    )
    parser.add_argument(
        "--excel_name",
        default=excel_name_default,
        help="Name of the Excel file (without extension).",
    )
    parser.add_argument(
        "--excel_format",
        default=".xlsx",
        help="Format of the Excel file (e.g., .xlsx, .xls).",
    )
    parser.add_argument(
        "--output_path",
        default=output_path_default,
        help="Path to the output directory for IDS files.",
    )

    args = parser.parse_args()
    
    if args == parser.parse_args([]):  
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
            raise Exception(f'{meta} is not defined in the Excel')
    
    sheet_name = data_dict['Sheet name']
    ifc_version = data_dict['IFC version'].replace(' ','').replace(',','|')
    separate_by = data_dict['File separators'].replace(' ','').split(',') if 'File separators' in data_dict else []
    skipped_rows = data_dict['Skipped rows'] if 'Skipped rows' in data_dict else 0
    is_entity_based_app = True if 'Entity-based applicability' in data_dict and data_dict['Entity-based applicability'].lower() == 'yes' else False
    
    return sheet_name, ifc_version, separate_by, skipped_rows, is_entity_based_app, data_dict

def process_excel_data(excel_path, excel_name, excel_format, sheet_name, separate_by, skipped_rows, ifc_version, is_entity_based_app):
    full_excel_path = os.path.join(excel_path, excel_name + excel_format)
    
    # Convert Excel data to specifications
    excel_data = excel_to_spec_list(full_excel_path, sheet_name, separate_by, skipped_rows, ifc_version, is_entity_based_app)
    
    # Separate specifications by general data
    separated_excel_data = separate_specs_by_generaldata(excel_data, separate_by)
    
    return separated_excel_data

def create_ids_files(separated_excel_data, data_dict, output_path, excel_name, sheet_name):
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

def main():
    args = parse_arguments()
    
    excel_path = args.excel_path
    excel_name = args.excel_name
    excel_format = args.excel_format
    output_path = args.output_path
    
    print(f"Excel file path: {os.path.join(excel_path, excel_name + excel_format)}") 
    print(f"Output path: {output_path}")
    print()
    
    # Extract metadata
    sheet_name, ifc_version, separate_by, skipped_rows, is_entity_based_app, data_dict = get_metadata(excel_path, excel_name, excel_format, 'IDS4ALL')
    
    # Process Excel data
    separated_excel_data = process_excel_data(excel_path, excel_name, excel_format, sheet_name, separate_by, skipped_rows, ifc_version, is_entity_based_app)
    
    # Create IDS files
    create_ids_files(separated_excel_data, data_dict, output_path, excel_name, sheet_name)

if __name__ == "__main__":
    main()
