# IDS4ALL converter

This IDS converter generates IDS files from information defined in tabular form. In current practice, it is common for information requirements to be defined in Excel files. The aim is to use a structure of the Excel file that is similar to existing structures for information requirements to facilitate the adaption to IDS. Therefore, our approach was to start with the classical structure and extend it to support the full functionality of the IDS, but still allow the use of the established structure. The structure allows one to start with a simple property assignment to an entity with a few columns and gradually increase the information density. The use and order of the defined columns are flexible. Only the names of the columns are predefined and must be adhered to. This gives the user great flexibility and allows existing tables to be used with only minor adjustments.

### Project structure

1. Folder [Example-IDS-files](Example-IDS-files): Contains the corresponding IDS files from the template and example Excel files contained in the folder [Excel-files](Excel-files).
2. Folder [Excel-files](Excel-files): Contains two Excel-files  
    - IDS4ALL-Template.xlsx: This is a template file containing a user manual for the developed tabular Excel structure, an example sheet including all possible functionalities, and the input sheet. 
    - Requirements-openBIM-building-permit-of-the-City-of-Vienna.xlsx: This is an example file containing the information requirements of the openBIM building permit of the City of Vienna.
3. Folder [output](output): Default output path for generated IDS files
4. [IDS4ALL-main.py](IDS4ALL-main.py): Main script to generate IDS files
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

Try the project with our example Excel ./Excel-files/Requirements-openBIM-building-permit-of-the-City-of-Vienna.xlsx

``` python IDS4ALL-main.py ```

Resulting IDS is written to
```./output/Requirements-openBIM-building-permit-of-the-City-of-Vienna_Specifications.ids```


### Override the input and output paths:

The default input and output paths can be overwritten either directly in the main script IDS4ALL-main.py or via command line arguments.

#### Override the input and output paths in the script:

To override the input and output paths and the name of the used Excel file you can edit the default settings in lines 9-11 in the main script [IDS4ALL-main.py](IDS4ALL-main.py):
```
#Default settings
excel_path_default = "./Excel-files/"
excel_name_default = "Requirements-openBIM-building-permit-of-the-City-of-Vienna"
output_path_default = "./output/"
```

#### Override the input and output paths via command line arguments

``` python IDS4ALL-main.py -h ```
```
usage: IDS4ALL-main.py [-h] [--excel_path EXCEL_PATH] [--excel_name EXCEL_NAME] [--excel_format EXCEL_FORMAT] [--output_path OUTPUT_PATH]

Generate IDS files from Excel data.

options:
  -h, --help            show this help message and exit
  --excel_path EXCEL_PATH
                        Path to the directory containing the Excel file.
  --excel_name EXCEL_NAME
                        Name of the Excel file (without extension).
  --excel_format EXCEL_FORMAT
                        Format of the Excel file (e.g., .xlsx, .xls).
  --output_path OUTPUT_PATH
                        Path to the output directory for IDS files.
```
