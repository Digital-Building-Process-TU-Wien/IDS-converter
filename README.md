# IDS4ALL converter

This IDS converter generates IDS files from information defined in tabular form. In current practice, it is common for information requirements to be defined in Excel files. The aim is to use a structure of the Excel file that is similar to existing structures for information requirements to facilitate the adaption to IDS. Therefore, our approach was to start with the classical structure and extend it to support the full functionality of the IDS, but still allow the use of the established structure. The structure allows one to start with a simple property assignment to an entity with a few columns and gradually increase the information density. The use and order of the defined columns are flexible. Only the names of the columns are predefined and must be adhered to. This gives the user great flexibility and allows existing tables to be used with only minor adjustments.

### Project structure

1. Folder [Example-IDS-files](Example-IDS-files): Contains the corresponding IDS files from the template and example Excel files contained in the folder [Excel-files](Excel-files).
2. Folder [Excel-files](Excel-files): Contains two Excel-files  
    - IDS4ALL-Template.xlsx: This is a template file containing a user manual for the developed tabular Excel structure, an example sheet including all possible functionalities, and the input sheet. 
    - Requirements-openBIM-building-permit-of-the-City-of-Vienna.xlsx: This is an example file containing the information requirements of the openBIM building permit of the City of Vienna.
3. Folder [output](output): Default output path for generated IDS or Excel files
4. [IDS4ALL-main.py](IDS4ALL-main.py): Main script to generate IDS or Excel files
5. [custom_functions.py](custom_functions.py): Additional module containing custom functions
6. [requirements.txt](requirements.txt): Text file containing the required dependencies

All data files are licensed under CC BY 4.0, all software files are licensed under MIT License.

## Donation

Our principle: sharing and distributing research is crucial to learning from each other and moving forward together. If you would like to support our open source research, we greatly appreciate donations.

Technische Universität Wien/Spendenkonto

IBAN: AT19 3200 0200 0061 1228

BIC/SWFT: RLNWATWW

Purpose of transfer: GEV011020ZFT S235030-000 - Digital Building Process

## Installation

1. Clone or download this repository
2. Install Python (compatible versions: 3.10, 3.11, 3.12)
3. Open Terminal in the project folder
4. Create a venv ``` python3 -m venv venv ``` or ``` python -m venv venv ```
5. Activate venv, Linux: ```source venv/bin/activate``` Windows: ``` venv\Scripts\activate ```
6. Install all dependencies: ```pip install -r requirements.txt```

## Run the project

### Default usage (Excel to IDS)

Try the project with our example Excel file:

``` python IDS4ALL-main.py ```

This will convert the default Excel file (./Excel-files/Requirements-openBIM-building-permit-of-the-City-of-Vienna-extended.xlsx) to IDS and write the result to ./output/.

### Convert IDS to Excel

To convert an IDS file to Excel format:

``` python IDS4ALL-main.py --input_format .ids --input_path ./Example-IDS-files/ --input_name "IDS4ALL-Template_Example specifications" ```

### Override the input and output paths

The default input and output paths can be overwritten via command line arguments.

#### Command-line arguments:

``` python IDS4ALL-main.py -h ```

``` 
usage: IDS4ALL-main.py [-h] [--input_path INPUT_PATH] [--input_name INPUT_NAME] [--input_format INPUT_FORMAT] [--output_path OUTPUT_PATH] 
                        [--excel_path EXCEL_PATH] [--excel_name EXCEL_NAME] [--excel_format EXCEL_FORMAT]
                        [--ids_path IDS_PATH] [--ids_name IDS_NAME] [--ids_format IDS_FORMAT]

Convert between Excel and IDS formats.

options:
  -h, --help            show this help message and exit
  --input_path INPUT_PATH
                        Path to the directory containing the input file.
  --input_name INPUT_NAME
                        Name of the input file (without extension).
  --input_format INPUT_FORMAT
                        Format of the input file (.xlsx, .xls, or .ids). Auto-detected from file extension if not provided.
  --output_path OUTPUT_PATH
                        Path to the output directory for generated files.
  --excel_path EXCEL_PATH
                        (Deprecated: use --input_path) Path to the directory containing the Excel file.
  --excel_name EXCEL_NAME
                        (Deprecated: use --input_name) Name of the Excel file (without extension).
  --excel_format EXCEL_FORMAT
                        (Deprecated: use --input_format) Format of the Excel file.
  --ids_path IDS_PATH   (Deprecated: use --input_path) Path to the directory containing the IDS file.
  --ids_name IDS_NAME
                        (Deprecated: use --input_name) Name of the IDS file (without extension).
  --ids_format IDS_FORMAT
                        (Deprecated: use --input_format) Format of the IDS file.
```

#### Examples

**Excel to IDS (default):**
```
python IDS4ALL-main.py --input_path ./Excel-files/ --input_name "IDS4ALL-Template" --input_format .xlsx --output_path ./output/
```

**IDS to Excel:**
```
python IDS4ALL-main.py --input_path ./Example-IDS-files/ --input_name "IDS4ALL-Template_Example specifications" --input_format .ids --output_path ./output/
```

**Using legacy arguments (Excel to IDS):**
```
python IDS4ALL-main.py --excel_path ./Excel-files/ --excel_name "IDS4ALL-Template" --excel_format .xlsx --output_path ./output/
```

**Using legacy arguments (IDS to Excel):**
```
python IDS4ALL-main.py --ids_path ./Example-IDS-files/ --ids_name "IDS4ALL-Template_Example specifications" --ids_format .ids --output_path ./output/
```

#### Behavior

- If `--input_format` is not specified, the converter attempts to auto-detect the format from the file extension. If neither is available, it defaults to `.xlsx` (Excel to IDS).
- An unknown file format will result in a clear error message.
- The converter automatically creates the output directory if it does not exist.
- Legacy command-line arguments (`--excel_*` and `--ids_*`) are still supported for backward compatibility.
