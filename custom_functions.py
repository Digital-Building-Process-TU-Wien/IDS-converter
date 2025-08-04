import pandas as pd
from xml.etree import ElementTree as ET
from deepdiff import DeepDiff
from deepmerge import Merger
import copy
from ifctester import ids
import itertools

STRING_ENTITY = 'Entity'
STRING_PREDEFINEDTYPE = 'PredefinedType'
STRING_PROPERTY = 'Property'
STRING_PROPERTYSET = 'PropertySet'
STRING_PROPERTYDATATYPE = 'PropertyDatatype'
STRING_PROPERTYVALUE = 'PropertyValue'
STRING_PROPERTYURI = 'PropertyURI'
STRING_MATERIAL = 'Material'
STRING_MATERIALURI = 'MaterialURI'
STRING_ATTRIBUTE = 'Attribute'
STRING_ATTRIBUTEVALUE = 'AttributeValue'
STRING_CLASSIFICATION = 'Classification'
STRING_CLASSIFICATIONSYSTEM = 'ClassificationSystem'
STRING_CLASSIFICATIONURI = 'ClassificationURI'
STRING_PARTOFRELATION = 'PartOfRelation'
STRING_PARTOFENTITY = 'PartOfEntity'
STRING_PARTOFPREDEFINEDTYPE = 'PartOfPredefinedType'
STRING_DESCRIPTION = 'Description'
STRING_SPECIFICATIONNAME = 'SpecificationName'
STRING_SPECIFICATIONCARDINALITY = 'SpecificationCardinality'
STRING_SPECIFICATIONIFCVERSION = 'SpecificationIfcVersion'
STRING_REQUIREMENTCARDINALITY = 'Cardinality'
KEYWORD_NONE = '_none_'
KEYWORD_MISSING = '_MISSING_'

entity_description_dict = {}
property_description_dict = {}

def excel_to_spec_list(EXCEL_PATH, sheet_name, separate_by, skipped_rows, ifc_versions, is_entity_based_app):
    '''Parses excel data from a given file path and sheet name into a list of specifications.
    Each specification is represented as dictionary with a given applicability, requirements, specification data, and general data

    :param EXCEL_PATH: Path to an excel file 
    :type EXCEL_PATH: str
    :param sheet_name: Name of the relevant sheet 
    :type sheet_name: str
    :param separate_by: List of general data for which specifications must be seperated
    :type separate_by: list
    :param skipped_rows: Number of skipped rows at the top of the sheet 
    :type skipped_rows: int
    :param ifc_versions: String specifying the default ifc versions for all specifications (are overwritten if the specificationIfcsVersion column is used) 
    :type ifc_versions: str
    :param is_entity_based_app: Boolean specifying if the applicability should be generated only entity-based (with predefiend types) 
    :type is_entity_based_app: boolean       
    :return: List of specifications with applicability, requirements, specification data and general data
    :rtype: list
    '''
    ###Import data
    all_columns = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, skiprows=skipped_rows, nrows=0).columns.tolist()
    dup = {x[:-2] for x in all_columns if x[-2:] == '.1'}
    if dup:
        raise Exception('Column names must be unique. The following column names occur multiple times: ' + ', '.join(map(str,dup)))
    applicability_data = []
    requirements_data = []
    generaldata = []
    specification_data = []

    #Organise all relevant available column names according to the facets
    prefix = 'A.'

    if is_entity_based_app:
        removed_columns = [prefix+STRING_PROPERTY,prefix+STRING_PROPERTYSET,prefix+STRING_PROPERTYVALUE,prefix+STRING_PROPERTYDATATYPE,
                           prefix+STRING_MATERIAL,
                           prefix+STRING_CLASSIFICATION,prefix+STRING_CLASSIFICATIONSYSTEM,
                           prefix+STRING_PARTOFENTITY,prefix+STRING_PARTOFRELATION]
        all_columns = [x for x in all_columns if x not in removed_columns]

    cols_app_entity = [col for col in [prefix+STRING_ENTITY] if col in all_columns]
    cols_app_property = [col for col in [prefix+STRING_PROPERTY,prefix+STRING_PROPERTYSET,prefix+STRING_PROPERTYVALUE,prefix+STRING_PROPERTYDATATYPE] if col in all_columns]
    cols_app_material = [col for col in [prefix+STRING_MATERIAL] if col in all_columns]
    cols_app_attribute = [col for col in [prefix+STRING_ATTRIBUTE,prefix+STRING_ATTRIBUTEVALUE] if col in all_columns]
    cols_app_classification = [col for col in [prefix+STRING_CLASSIFICATION,prefix+STRING_CLASSIFICATIONSYSTEM] if col in all_columns]
    cols_app_partOf = [col for col in [prefix+STRING_PARTOFENTITY,prefix+STRING_PARTOFRELATION] if col in all_columns]
    prefix = 'R.'
    cols_req_entity = [col for col in [prefix+STRING_ENTITY,prefix+STRING_DESCRIPTION+'.Entity'] if col in all_columns]
    cols_req_property = [col for col in [prefix+STRING_PROPERTY,prefix+STRING_PROPERTYSET,prefix+STRING_PROPERTYVALUE,prefix+STRING_PROPERTYDATATYPE,prefix+STRING_PROPERTYURI,prefix+STRING_REQUIREMENTCARDINALITY,prefix+STRING_DESCRIPTION+'.Property'] if col in all_columns]
    cols_req_material = [col for col in [prefix+STRING_MATERIAL,STRING_MATERIALURI,prefix+STRING_REQUIREMENTCARDINALITY] if col in all_columns]
    cols_req_attribute = [col for col in [prefix+STRING_ATTRIBUTE,prefix+STRING_ATTRIBUTEVALUE,prefix+STRING_REQUIREMENTCARDINALITY] if col in all_columns]
    cols_req_classification = [col for col in [prefix+STRING_CLASSIFICATION,prefix+STRING_CLASSIFICATIONSYSTEM,prefix+STRING_CLASSIFICATIONURI,prefix+STRING_REQUIREMENTCARDINALITY] if col in all_columns]
    cols_req_partOf = [col for col in [prefix+STRING_PARTOFENTITY,prefix+STRING_PARTOFRELATION,prefix+STRING_REQUIREMENTCARDINALITY] if col in all_columns]
    
    cols_general = [col for col in ['Phase','Role','Usecase'] if col in all_columns]

    cols_specification = [col for col in [STRING_SPECIFICATIONCARDINALITY,STRING_SPECIFICATIONIFCVERSION] if col in all_columns]

    #Store all relevant column names of the used file
    relevant_columns = []
    relevant_columns.extend(cols_app_entity)
    relevant_columns.extend(cols_app_property)
    relevant_columns.extend(cols_app_material)
    relevant_columns.extend(cols_app_attribute)
    relevant_columns.extend(cols_app_classification)
    relevant_columns.extend(cols_app_partOf)
    relevant_columns.extend(cols_req_entity)
    relevant_columns.extend(cols_req_property)
    relevant_columns.extend(cols_req_material)
    relevant_columns.extend(cols_req_attribute)
    relevant_columns.extend(cols_req_classification)
    relevant_columns.extend(cols_req_partOf)
    relevant_columns.extend(cols_general)
    relevant_columns.extend(cols_specification)
    relevant_columns = list(set(relevant_columns))
    
    ##Import the relevant columns and merge rows with the same applicability into one row
    if relevant_columns:
        #Import all relevant columns
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, skiprows=skipped_rows, usecols=relevant_columns)

        #Fill empty requirement cardinality with default value
        if prefix+STRING_REQUIREMENTCARDINALITY in df.columns:
            df[prefix+STRING_REQUIREMENTCARDINALITY] = df[prefix+STRING_REQUIREMENTCARDINALITY].fillna("required")

        #Fill empty specification ifc version with default version from IDS4ALL sheet
        spec_ifc_version_col_used = False
        if STRING_SPECIFICATIONIFCVERSION not in df.columns:
            df[STRING_SPECIFICATIONIFCVERSION] = ifc_versions
            relevant_columns.append(STRING_SPECIFICATIONIFCVERSION)
            cols_specification.append(STRING_SPECIFICATIONIFCVERSION)
        else:
            df[STRING_SPECIFICATIONIFCVERSION] = df[STRING_SPECIFICATIONIFCVERSION].fillna(ifc_versions)
            spec_ifc_version_col_used = True

        #Fill NaN values with a placeholder
        df_filled = df.fillna(KEYWORD_MISSING)

        #Step 1: Merge rows of the columns STRING_PROPERTYVALUE & STRING_ATTRIBUTEVALUE if the values in the other columns are identical
        relevant_columns_copy = relevant_columns.copy()
        merge_columns_1 = [prefix+STRING_PROPERTYVALUE,
                        prefix+STRING_ATTRIBUTEVALUE]
        removed_items = []
        for item in merge_columns_1:
            if item in relevant_columns_copy:
                relevant_columns_copy.remove(item)
                removed_items.append(item)

        df_step1 = df_filled.groupby(relevant_columns_copy, dropna=False, sort=False).agg({
            col: lambda x: '|'.join(map(str, x.dropna())) for col in removed_items
            }).reset_index()

        #Step 2: Merge all requirement parameters and general parameters not in seperate_by,
        #keeping values of STRING_PROPERTYVALUE & STRING_ATTRIBUTEVALUE as sub-lists
        merge_columns_2 = []
        merge_columns_2.extend(cols_req_entity)
        merge_columns_2.extend(cols_req_property)
        merge_columns_2.extend(cols_req_material)
        merge_columns_2.extend(cols_req_attribute)
        merge_columns_2.extend(cols_req_classification)
        merge_columns_2.extend(cols_req_partOf)
        merge_columns_2.extend([item for item in cols_general if item not in separate_by])
        merge_columns_2 = list(set(merge_columns_2))
        for item in merge_columns_2:
            if item in relevant_columns_copy:
                relevant_columns_copy.remove(item)
                removed_items.append(item)
        
        df_final = df_step1.groupby(relevant_columns_copy, as_index=False, sort=False).agg({
            col: lambda x: list(x) for col in removed_items  # Ensure sublists for dynamic value columns
            })

        #Create individual dataframes for each facet and store them structured in applicability, requirements, general, and specification
        if cols_app_entity:
            applicability_data.append(df_final[cols_app_entity])
        if cols_app_property:
            applicability_data.append(df_final[cols_app_property])
        if cols_app_material:
            applicability_data.append(df_final[cols_app_material])
        if cols_app_attribute:
            applicability_data.append(df_final[cols_app_attribute])
        if cols_app_classification:
            applicability_data.append(df_final[cols_app_classification])
        if cols_app_partOf:
            applicability_data.append(df_final[cols_app_partOf])
        if cols_req_entity:
            requirements_data.append(df_final[cols_req_entity])
        if cols_req_property:
            requirements_data.append(df_final[cols_req_property])
        if cols_req_material:
            requirements_data.append(df_final[cols_req_material])
        if cols_req_attribute:
            requirements_data.append(df_final[cols_req_attribute])
        if cols_req_classification:
            requirements_data.append(df_final[cols_req_classification])
        if cols_req_partOf:
            requirements_data.append(df_final[cols_req_partOf])
        if cols_general:
            generaldata.append(df_final[cols_general])
        if cols_specification:
            specification_data.append(df_final[cols_specification])

    ###Transform each dataframe row into a dictionary
    specs_list = []
    for i in range(df_final.index.size):
        #General data
        generaldata_dict = {}
        if generaldata:
            row_generaldata = generaldata[0].iloc[i]
            generaldata_dict = pandas_row_to_dict(row_generaldata)

        #Specification data
        specification_data_dict = {}
        if specification_data:
            row_specification_data = specification_data[0].iloc[i]
            specification_data_dict = pandas_row_to_dict(row_specification_data)
            if STRING_SPECIFICATIONCARDINALITY in specification_data_dict:
                if specification_data_dict[STRING_SPECIFICATIONCARDINALITY][0].lower() not in ['required', 'prohibited']:
                    specification_data_dict.pop(STRING_SPECIFICATIONCARDINALITY)
            if STRING_SPECIFICATIONIFCVERSION in specification_data_dict:
                for ifc_version in specification_data_dict[STRING_SPECIFICATIONIFCVERSION]:
                    if ifc_version.upper() not in ['IFC2X3','IFC4','IFC4X3_ADD2']:
                        raise Exception('Invalid IFC version used: ' + ifc_version)
       
        #Applicability data
        app_list = []
        for applicability_facet_df in applicability_data:
            row_app = applicability_facet_df.iloc[i]
            dict_list_app = pandas_row_to_dict_list(row_app)
            for dict_item in dict_list_app:
                    app_dict = split_OR_AND_values(dict_item, is_entity_based_app)
                    #Rearrange 'AND' values of the applicability into individual facet dictionaries
                    app_list.extend(split_AND_values_to_individual_facet_dicts(app_dict))

        #Generate possible combinations for all or_values in the applicability and create individual specifications for each.
        #This is necessary for assigning generally applicable values to more specific specifications.
        #Only if specification cardinality is not "required", because in this case the logic would be lost when seperated
        if STRING_SPECIFICATIONCARDINALITY in specification_data_dict and specification_data_dict[STRING_SPECIFICATIONCARDINALITY][0].lower() == 'required':
            app_lists = [app_list]
        else:
            app_lists = generate_combinations(app_list)
        for app_list in app_lists:
            for app_dict in app_list:
                #Separate Entity.PredefinedType into two entries
                separate_dict_value(app_dict, STRING_ENTITY, STRING_PREDEFINEDTYPE,'.')
                separate_dict_value(app_dict, STRING_PARTOFENTITY, STRING_PARTOFPREDEFINEDTYPE,'.')

            #Check if a specification with this general data and applicability already exists
            req_list = []
            diff, diff_generaldata, diff_specification_data, req_list = compare_previous_generaldata_and_applicability(specs_list, generaldata_dict, specification_data_dict, app_list, separate_by)
                    
            #Requirements data
            for requirement_facet_df in requirements_data:
                row_req = requirement_facet_df.iloc[i]
                dict_list_req = pandas_row_to_dict_list(row_req)
                for dict_item in dict_list_req:
                    req_dict = split_OR_AND_values(dict_item, is_entity_based_app)
                    #check if dict only includes the cardinality (invalid); if so, take empty dict.
                    #Can occur since requirement cardinality column is applied to several facets from which some might be empty in current row
                    if len(req_dict.keys()) == 1 and STRING_REQUIREMENTCARDINALITY in req_dict:
                        req_dict = {}
                    #Rearrange 'AND' values of the requirements into individual facet dictionaries
                    req_dict_list_arranged = split_AND_values_to_individual_facet_dicts(req_dict)
                    
                    for req_dict_arranged in req_dict_list_arranged:
                        #Check whether complex restrictions are included in 'OR' values (this is not possible in IDS)
                        for key in req_dict_arranged:
                            if len(req_dict_arranged[key]) > 1:
                                j = 0
                                while j < len(req_dict_arranged[key]):
                                    value = req_dict_arranged[key][j]
                                    if is_complex_restriction(value):
                                        raise Exception('Complex restrictions (pattern=; \\<=; \\<; \\>=; \\>; length=; length<=; length>=) cannot be part of an enumeration (list of "or"-values)')
                                    else: j += 1
                        
                        #Extract descriptions for properties
                        #Necessary to add property description to each occurance in the IDS even if it is only one in the Excel
                        #Not necessary for entity descriptions
                        if STRING_DESCRIPTION in req_dict_arranged:
                            if STRING_PROPERTY in req_dict_arranged:
                                property_description_dict[req_dict_arranged[STRING_PROPERTY][0]] = req_dict_arranged[STRING_DESCRIPTION][0]
                        
                        #Separate Entity.PredefinedType into two entries
                        separate_dict_value(req_dict_arranged, STRING_ENTITY, STRING_PREDEFINEDTYPE,'.')
                        separate_dict_value(req_dict_arranged, STRING_PARTOFENTITY, STRING_PARTOFPREDEFINEDTYPE,'.')
                        
                        #If the same requirement (except for the values) already exists, merge the values.
                        #Otherwise, add new requirement
                        diff_req = True
                        for j in range(len(req_list)):
                            diff_req = compare_and_merge_requirement_dicts(req_list[j], req_dict_arranged, [STRING_ATTRIBUTEVALUE,STRING_PROPERTYVALUE], False)
                            if not diff_req: break
                        if diff_req: req_list.append(req_dict_arranged)

            #if it is a new spec (with different applicability, generaldata, or specificationdata), create a new specification
            if diff or diff_generaldata or diff_specification_data:
                if app_list or req_list or generaldata_dict:
                    spec_dict = {}
                    spec_dict['app'] = app_list
                    spec_dict['req'] = req_list
                    spec_dict['general'] = copy.deepcopy(generaldata_dict)
                    spec_dict['spec'] = copy.deepcopy(specification_data_dict)
                    specs_list.append(spec_dict)

    #organise the specifications according to the ifc versions
    #if one specification refers to a subset of ifc versions of another specification with the same applicability and general data,
    #the subset of ifc versions is extracted from the more general specification and the requirements are included into the specific specification.
    #this allows to merge specifications with the same applicability, general data, and ifc versions
    if spec_ifc_version_col_used:
        for i in range(len(specs_list)):
            specI = specs_list[i]
            specI_Ifc_versions = specI['spec'][STRING_SPECIFICATIONIFCVERSION]
            if not specI_Ifc_versions: continue
            for j in range(i+1,len(specs_list)):
                specI_Ifc_versions = specI['spec'][STRING_SPECIFICATIONIFCVERSION]
                if not specI_Ifc_versions: break
                specJ = specs_list[j]
                specJ_Ifc_versions = specJ['spec'][STRING_SPECIFICATIONIFCVERSION]
                if not specJ_Ifc_versions: continue
                #if all ifc versions of specI are in specJ, specI_Ifc_versions is a subset of specJ_Ifc_versions
                #Then all requirements of specJ also apply to specI.
                if (all(elem in specJ_Ifc_versions for elem in specI_Ifc_versions)):
                    structure_specifications_by_Ifc_versions(specI, specJ, specI_Ifc_versions, specJ_Ifc_versions, separate_by)
                #if all ifc versions of specJ are in specI, specJ_Ifc_versions is a subset of specI_Ifc_versions
                #Then all requirements of specI also apply to specJ.
                elif (all(elem in specI_Ifc_versions for elem in specJ_Ifc_versions)):
                    structure_specifications_by_Ifc_versions(specJ, specI, specJ_Ifc_versions, specI_Ifc_versions, separate_by)
                
        #remove specifications with empty ifc versions (might be created during the re-structuring)
        specs_list = [spec for spec in specs_list if spec['spec'][STRING_SPECIFICATIONIFCVERSION]]

    #add requirements of specific specifications to more general specifications
    add_values_to_general_specs(specs_list,separate_by)

    return specs_list

def load_columns(EXCEL_PATH, sheet_name, skipped_rows, columns_to_import, all_columns, dataframe_list):
    '''Loads the columns of columns_to_import that are present in the excel sheet and parses them into a pandas dataframe.
    The available columns in the excel sheet are given by the all_columns list 

    :param EXCEL_PATH: Path to an excel file 
    :type EXCEL_PATH: str
    :param sheet_name: Name of the relevant sheet 
    :type sheet_name: str
    :param skipped_rows: Number of skipped rows at the top of the sheet 
    :type skipped_rows: int
    :param columns_to_import: List of column names to import. 
    :type columns_to_import: List
    :param all_columns: List of all column names in the sheet. 
    :type all_columns: List
    :param all_columns: List containing imported data frames
    :type all_columns: List
    :return: Pandas dataframe containing the data of the imported columns
    :rtype: Pandas dataframe
    '''
    available_columns = [col for col in columns_to_import if col in all_columns]
    if available_columns:
        dataframe_list.append(pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, skiprows=skipped_rows, usecols=available_columns))

def pandas_row_to_dict(current_row):
    '''Stores values of the current row of a pandas dataframe into a dictionary using the column names as keys.
    Uses '\\&' and '|' as delimiter but does not differentiate in their meaning. Stores the values in a list and assigns it to a key in the dictionary.
    (key = [value, value, value])

    :param current_row: Pandas row
    :type current_row: Pandas row
    :return: Dictionary using the column names as keys and the row values as values (key = [value, value, value])
    :rtype: dict
    '''
    new_dict = {}
    for column_name, column_data in current_row.items():
        #Cut characters after the first '.'. If duplicate column names exist, pandas numbers them with name.1, name.2 ...
        column_name = column_name.split('.')[0]
        if isinstance(column_data,list):
            column_data = '|'.join(column_data)
        #Only use strings and numbers not null        
        if isinstance(column_data,str): #or not np.isnan(column_data):
            if column_data != KEYWORD_NONE and column_data != KEYWORD_MISSING:
                column_data = str(column_data).strip()
                #Use delimiters to distinguish between different values (for general data, the delimiters have the same meaning)
                or_values = column_data.replace('\\&','|').split('|')
                #Omit KEYWORD_MISSING values
                or_values_cleaned = []
                values_list = []
                for or_value in or_values:
                    if or_value != KEYWORD_MISSING and or_value not in or_values_cleaned:
                        or_values_cleaned.append(or_value)
                if or_values_cleaned:
                    values_list.extend(or_values_cleaned)
                #If or values except KEYWORD_MISSING exist, add them to the dictionary
                if values_list:
                    new_dict[column_name] = values_list
    return new_dict

def pandas_row_to_dict_list(current_row):
    '''Creates individual dictionaries for each facet in the current row of the given pandas dataframe using the column names as keys.
    Then removes dictionaries if they are subsets of another dictionary.

    :param current_row: Pandas row
    :type current_row: Pandas row
    :return: List of individual dictionaries for each facet
    :rtype: list
    '''
    list_columns = current_row.index  # Get all column names
    list_values = [current_row[col] if isinstance(current_row[col], list) else [current_row[col]] for col in list_columns]  # Ensure lists

    num_items = max(len(values) for values in list_values)  # Determine max length of lists

    # Create dictionaries while handling non-list values and skipping KEYWORD_MISSING
    dict_list = [
        {
            col: values[i] if i < len(values) else values[0]  # Use existing value or repeat if shorter
            for col, values in zip(list_columns, list_values)
            if values[i] != KEYWORD_MISSING  # Skip KEYWORD_MISSING values
        }
        for i in range(num_items)
    ]

    #Remove dictionaries from the list if they are subsets of another dictionary.
    filtered_list = []
    for d1 in dict_list:
        if not any(d1.items() <= d2.items() and d1 != d2 for d2 in dict_list):
            filtered_list.append(d1)
    
    return filtered_list

def split_OR_AND_values(input_dict, is_entity_based_app):
    '''Transforms the values of the dictionary from strings to list of strings using delimiters.
    The delimiter '\\&' is used for 'AND' values and '|' is used for 'OR' values.
    The lists of 'OR' values are nested in lists of 'AND' values and then assigned to their original key in a new dictionary.
    'AND' values contain lists of 'OR' values:
    key: [[OR_value, OR_value],[OR_value, OR_value, OR_value],[OR_value]]

    :param input_dict: Dictionary containing values as strings
    :type input_dict: dict
    :param is_entity_based_app: Boolean specifying if the applicability should be generated only entity-based (with predefiend types)
    :type is_entity_based_app: boolean
    :return: Dictionary containing nested lists for 'AND' and 'OR' values (key: [[OR_value, OR_value],[OR_value, OR_value, OR_value],[OR_value]])
    :rtype: dict
    '''
    new_dict = {}
    number_of_and_values = 0
    if input_dict:
        #Rearrange 'AND' values into individual facet dictionaries
        keylist = input_dict.keys()
        for key in keylist:
            value = input_dict[key]
            key = key.split('.')[1]
            values_list = []
            if isinstance(value,str):
                if value != KEYWORD_NONE and value != KEYWORD_MISSING:
                    value = str(value).strip()
                    #Change relevant values to uppercase
                    if key in [STRING_PROPERTYDATATYPE,STRING_PARTOFRELATION]:
                        value = value.upper()
                    
                    #Use delimiters to distinguish between different 'AND' values
                    and_values = value.split('\\&')
                    #Check whether all entries of this facet have the same number of 'AND' values
                    if number_of_and_values == 0:
                        number_of_and_values = len(and_values)
                    if len(and_values) != number_of_and_values:
                        #Special case requirement cardinality. Since the requirement cardinality is included automatically if not in the excel file,
                        #if the existing columns use 'and' values the cardinality has no and values. Then the one value is applied to all and values.
                        if key == STRING_REQUIREMENTCARDINALITY and len(and_values) == 1:
                            and_values = and_values*number_of_and_values
                        else:
                            raise Exception('Number of AND values (seperated by \\&) is invalid for ' + key + ' ' + value + '. All columns of one IDS facet require the same number of AND values. Required number of AND values: ' + str(number_of_and_values))
                    values_list = []
                    for and_value in and_values:
                        #Use delimiters to distinguish between different 'OR' values
                        or_values = None if and_value == '' else and_value.split('|')
                        if or_values == None: continue
                        #Check if complex restrictions were merged due to entity-based applicability. If so, delete them
                        if is_entity_based_app and len(or_values) > 1:
                            or_values = list(filter(lambda or_value: not is_complex_restriction(or_value), or_values))
                        #Omit KEYWORD_MISSING values
                        or_values_cleaned = []
                        for or_value in or_values:
                            if or_value != KEYWORD_MISSING:
                                or_values_cleaned.append(or_value)
                        if or_values_cleaned:
                            values_list.append(or_values_cleaned)

            elif isinstance(value,bool) or isinstance(value,int) or isinstance(value,float) or isinstance(value,complex):
                values_list.append([value])

            #If values except KEYWORD_MISSING exist, add them to the dictionary            
            if values_list:
                new_dict[key] = values_list

    return new_dict

def split_AND_values_to_individual_facet_dicts(input_dict):
    '''Creates and individual dictionary for each 'AND' value in the dictionary.
    The input dict contains lists of 'AND' values for each key. These are split in individual dictionaries including only one value for each key.

    :param input_dict: Dictionary containing lists of 'AND' values for each key
    :type input_dict: dict
    :return: List of individual dictionaries for each 'AND' value (each facet)
    :rtype: list
    '''
    facet_list = []
    if input_dict:
        #Rearrange 'AND' values into individual facet dictionaries
        keylist = input_dict.keys()
        for j in range(len(input_dict[list(keylist)[0]])):
            facet_dict = {}
            for key in keylist:
                facet_dict[key] = input_dict[key][j]
            facet_list.append(facet_dict)
    return facet_list

def separate_dict_value(current_dict, combined_key, second_key, delimiter):
    '''Separates the value of 'combined_key' in the current dict into two values using a delimiter.
    The part before the delimiter is stored as value for the 'combined_key' and the part after the delimiter as value for the second_key.

    :param current_dict: Used dictionary
    :type current_dict: dict
    :param combined_key: initial key for the combined value (and key for the first part of the combined value)
    :type combined_key: str
    :param second_key: new key for the second part of the combined value
    :type second_key: str
    :param delimiter: delimiter for the separation
    :type delimiter: str
    :return: None
    '''
    if combined_key in current_dict:
        or_values = current_dict[combined_key]
        or_values_first_key = []
        or_values_second_key = []
        for or_value in or_values:
            separate_values = or_value.split(delimiter)
            if combined_key == STRING_ENTITY or combined_key == STRING_PARTOFENTITY: separate_values[0] = separate_values[0].upper()
            or_values_first_key.append(separate_values[0])
            #Only append value to second key if it exists
            if len(separate_values) == 2:
                or_values_second_key.append(separate_values[1])

        if len(set(or_values_first_key)) > 1 and len(or_values_second_key) > 0:
            raise Exception("PredefinedTypes cannot be used in an enumeration in a specification applicability with cardinality \"required\" because separating it into two different facet parameters causes the connection between the PredefinedType and the Entity to be lost. Affected Entity Facet: Entities: " + str(or_values_first_key) + " PredefinedTypes: " + str(or_values_second_key))                           
        current_dict[combined_key] = or_values_first_key
        if or_values_second_key: current_dict[second_key] = or_values_second_key

def generate_combinations(data):
    '''Generate all possible combinations of dictionaries with list values.
    This function takes a list of dictionaries, where each dictionary may 
    contain keys associated with lists of values. It produces all possible 
    combinations such that for each dictionary, if a key has multiple values 
    in its list, a separate dictionary is created for each value. The function 
    ensures that the structure of the input dictionaries is preserved, and 
    each resulting combination includes one value per list.

    :param data: List of dictionaries, where each key maps to a list of values.
    :type current_dict: list
    :return: List of lists, where each inner list represents a unique combination of dictionaries, preserving the input structure but with one value per list.
    :rtype: list
    '''
    # Collect all possible values for each dictionary in `data`
    options = []
    for item in data:
        item_variants = []
        # Create single-value dictionaries for each key in `item`
        for key, values in item.items():
            # If multiple values, create a separate dictionary for each value
            if values is not None and len(values) > 1:
                item_variants.append([{key: [value]} for value in values])
            else:
                # If only one value, use it as-is
                item_variants.append([{key: values}])
        # Combine the single-value dictionaries for each key in `item`
        options.append([dict(sum((list(d.items()) for d in combo), [])) for combo in itertools.product(*item_variants)])
    
    # Cartesian product of all item combinations
    combined_variants = [list(variant) for variant in itertools.product(*options)]
    return combined_variants

def compare_previous_generaldata_and_applicability(specs_list, generaldata_dict, specification_data_dict, app_list, separate_by):
    '''Checks if a specification with the same applicability and general data already exists.
    For the general data check only the general data speficied by separate_by is relevant.
    Other general data can be merged anyways.
    If a specification with the same applicability and general data already exists, the requirements of the new and previous are merged.

    :param specs_list: List of all previous specifications
    :type specs_list: list
    :param generaldata_dict: Dictionary containing lists of 'AND' values for each generaldata key
    :type generaldata_dict: dict
    :param specification_data_dict: Dictionary containing lists of 'AND' values for each specificationdata key
    :type specification_data_dict: dict
    :param app_list: List containing dictionaries for each facet in the applicability
    :type app_list: list
    :param separate_by: List of general data for which specifications must be seperated
    :type separate_by: list    
    :return: Boolean 'diff' specifying whether the applicability is different
    :rtype: bool
    :return: Boolean 'diff_generaldata' specifying whether the generaldata is different
    :rtype: list        
    :return: List containing the requriements
    :rtype: list
    '''
    req_list = []
    diff = True
    diff_generaldata = True
    diff_specification_data = True
    for j in range(len(specs_list)-1,-1,-1):
        prev_spec = specs_list[j]
        #check specification data
        diff_specification_data = DeepDiff(specification_data_dict, prev_spec['spec'], include_paths=[STRING_SPECIFICATIONCARDINALITY,STRING_SPECIFICATIONIFCVERSION])
        if not diff_specification_data:
            #check the general data specified by seperate_by, if it is not empty.
            if separate_by:
                diff_generaldata = DeepDiff(generaldata_dict, prev_spec['general'], include_paths=separate_by)
            else: diff_generaldata = False
            if not diff_generaldata:
                #check the applicability
                diff = app_list != prev_spec['app']
                if not diff:
                    diff = DeepDiff(app_list, prev_spec['app'])
                if not diff:
                    #merge the requirements
                    req_list = prev_spec['req']
                    prev_spec['general'] = merger.merge(prev_spec['general'], generaldata_dict)
                    break
    return diff, diff_generaldata, diff_specification_data, req_list

def structure_specifications_by_Ifc_versions(spec1, spec2, spec1_Ifc_versions, spec2_Ifc_versions, separate_by):
    '''Includes all requirements of spec2 in spec1 and deletes the ifc versions of spec1 from spec2,
    if the two specifications have equal applicability, general data and specification cardinality.
    If all ifc versions of spec1 are in spec2, spec1_Ifc_versions is a subset of spec2_Ifc_versions. Then all requirements of spec2 also apply to spec1.
    After the requirements are included in spec1, the ifc versions of spec1 are removed spec2.
    By this, the two specifications with overlapping ifc-versions are seperated to have one specification for each unique ifc version combination

    :param spec1: Specification with more specific ifc version definition
    :type spec1: dict
    :param spec2: Specification with more general ifc version definition
    :type spec2: dict
    :param spec1_Ifc_versions: List of all ifc versions spec1 applies to
    :type spec1_Ifc_versions: list
    :param spec2_Ifc_versions: List of all ifc versions spec2 applies to
    :type spec2_Ifc_versions: list
    :param separate_by: List of general data for which specifications must be seperated
    :type separate_by: list 
    '''
    #check specification data
    diff_specification_data = DeepDiff(spec1['spec'], spec2['spec'], include_paths=[STRING_SPECIFICATIONCARDINALITY])
    if not diff_specification_data:
        #check the general data specified by seperate_by, if it is not empty.
        if separate_by:
            diff_generaldata = DeepDiff(spec1['general'], spec2['general'], include_paths=separate_by)
        else: diff_generaldata = False
        if not diff_generaldata:
            #check the applicability
            diff = spec1['app'] != spec2['app']
            if not diff:
                diff = DeepDiff(spec1['app'], spec2['app'])
            if not diff:
                spec1['general'] = merger.merge(spec1['general'], spec2['general'])
                for req_dict2 in spec2['req']:
                    found = False
                    for req_dict1 in spec1['req']:
                        found = not compare_and_merge_requirement_dicts(req_dict1, req_dict2, [STRING_ENTITY,STRING_PREDEFINEDTYPE,STRING_ATTRIBUTEVALUE,STRING_PROPERTYVALUE], False, True)
                        if found == True: break
                        
                    if not found:
                        # Shallow copy to avoid mutating original item2 later
                        spec1['req'].append({k: v[:] if isinstance(v, list) else v for k, v in req_dict2.items()})
                #Since all requirements of spec2 are included in spec1, spec2 does not need to apply to the ifc versions of spec1 anymore
                spec2['spec'][STRING_SPECIFICATIONIFCVERSION] = list(set(spec2_Ifc_versions) - set(spec1_Ifc_versions))

def add_values_to_general_specs(specs_list, separate_by):
    '''Checks for all specification if a specification with a more general applicability exists.
    If so, the requirements of the more specific specification are added to the more general specification.

    :param specs_list: List of all specifications
    :type specs_list: list
    :param separate_by: List of general data for which specifications must be seperated
    :type separate_by: list 
    '''
    for i in range(len(specs_list)):
        for j in range(i+1,len(specs_list)):
            specI = specs_list[i]
            specJ = specs_list[j]
            
            #check if specification meta data of both specs is equal
            diff_specification_data = False
            specdataI = specI['spec']
            specdataJ = specJ['spec']
            diff_specification_data = DeepDiff(specdataI, specdataJ, ignore_order=True, include_paths=[STRING_SPECIFICATIONCARDINALITY,STRING_SPECIFICATIONIFCVERSION])

            #check if general data of both specs is equal
            if not diff_specification_data:
                diff_generaldata = False
                if separate_by:
                    generaldataI = specI['general']
                    generaldataJ = specJ['general']
                    diff_generaldata = DeepDiff(generaldataI, generaldataJ, ignore_order=True, include_paths=separate_by)
                
                #check if spec i has a more general applicability than j.
                if not diff_generaldata:
                    appI = specI['app']
                    appJ = specJ['app']
                    if len(appJ) > 0 and len(appI) > 0:
                        if STRING_ENTITY in appI[0] and STRING_ENTITY in appJ[0]:
                            if appI[0][STRING_ENTITY] != appJ[0][STRING_ENTITY]: continue
                        if STRING_PREDEFINEDTYPE in appI[0] and STRING_PREDEFINEDTYPE in appJ[0]:
                            if appI[0][STRING_PREDEFINEDTYPE] != appJ[0][STRING_PREDEFINEDTYPE]: continue
                    if len(appJ) > 1 and len(appI) > 1:
                        if STRING_ATTRIBUTEVALUE in appI[1] and STRING_ATTRIBUTEVALUE in appJ[1]:
                            if appI[1][STRING_ATTRIBUTEVALUE] != appJ[1][STRING_ATTRIBUTEVALUE]: continue

                    diff_app = DeepDiff(appI, appJ, ignore_order=True)
                    removed_dict_items_appJ = diff_app.get('dictionary_item_removed',[])
                    removed_iterable_items_appJ = diff_app.get('iterable_item_removed',[])
                    added_dict_items_appJ = diff_app.get('dictionary_item_added',[])
                    added_iterable_items_appJ = diff_app.get('iterable_item_added',[])
                    values_changed_appJ = diff_app.get('values_changed',[])

                    #if general data is equal and app more general in appI, compare/merge the requirements of appJ in appI
                    if not removed_dict_items_appJ and not removed_iterable_items_appJ and not values_changed_appJ:
                        #compare the requirements
                        reqI = specI['req']
                        reqJ = specJ['req']
                        for k in range(len(reqI)):
                            for l in range(len(reqJ)):
                                diff_req = compare_and_merge_requirement_dicts(reqI[k], reqJ[l], [STRING_ATTRIBUTEVALUE,STRING_PROPERTYVALUE], True)
                    
                    #if general data is equal and app more general in appJ, compare/merge the requirements of appI in appJ
                    if not added_dict_items_appJ and not added_iterable_items_appJ and not values_changed_appJ:
                        #compare the requirements
                        reqI = specI['req']
                        reqJ = specJ['req']
                        for k in range(len(reqI)):
                            for l in range(len(reqJ)):
                                diff_req = compare_and_merge_requirement_dicts(reqJ[l], reqI[k], [STRING_ATTRIBUTEVALUE,STRING_PROPERTYVALUE], True)                    

def compare_and_merge_requirement_dicts(req_dict1, req_dict2, not_compared_keys, merge_only_values, reverse=False):
    '''Checks whether the req_dict1 is a subset of req_dict2.
    This means, the old requirement must be a subset of the new requirement. The keys of req_dict1 must exist in req_dict2 and the values must be equal.
    An exception are lists for property values or attribute values. These are not compared. Here the values can be different because they are merged together, unless one contains a complex restriction. Complex restrictions cannot be merged and thus always indicate a difference.
    If req_dict1 is a subset of req_dict2, the information of req_dict2 is merged into req_dict1.
    
    :param req_dict1: dictionary 1 (is altered if dictionaries are subsets)
    :type req_dict1: dict
    :param req_dict2: dictionary 2 (is not altered)
    :type req_dict2: dict
    :param merge_only_values: Boolean defining that dicts should only be merged if all keys except Description exist in both dicts
    :type merge_only_values: boolean
    :param reverse: Boolean defining if the subset logic should be reversed. If true, req_dict2 must be a subset of req_dict1, but still req_dict2 is merged into req_dict1
    :type reverse: boolean
    :return: boolean specifying whether the dicts were different or not
    :rtype: boolean
    '''
    diff = False
    #boolean specifying whether the same type of facet is compared (e.g. property facet with property facet)
    same_facet = False

    #If property or attribute values are included and are a complex restriction, the requirements must not be merged
    #Complex restrictions cannot be in an enumeration
    for key in not_compared_keys:
        diff = True if key in req_dict1 and req_dict1[key] and is_complex_restriction(req_dict1[key][0]) else diff
        diff = True if key in req_dict2 and req_dict2[key] and is_complex_restriction(req_dict2[key][0]) else diff

    #Define which dict must be a subset of the other according to the reverse parameter
    if not reverse:        
        looped_dict = req_dict1
        reference_dict = req_dict2
    else:
        looped_dict = req_dict2
        reference_dict = req_dict1

    for key in looped_dict:
        if key == STRING_DESCRIPTION: continue
        #if only values should be merged, all keys except STRING_DESCRIPTION must also occur in both dicts
        if merge_only_values:
                keys1 = set(looped_dict.keys()) - set([STRING_DESCRIPTION])
                keys2 = set(reference_dict.keys()) - set([STRING_DESCRIPTION])
                if keys1 != keys2: 
                    diff = True
                    break
        #all keys in req_dict1 must be equal in req_dict2
        if key in reference_dict:
            same_facet = True
            #lists for property values and attribute values are not compared
            if key in not_compared_keys:
                continue
            elif req_dict1[key] != req_dict2[key]: diff = True
        else: diff = True
    if not diff and same_facet:
        req_dict1 = merger.merge(req_dict1, req_dict2)
        return False
    return True

def is_complex_restriction(value):
    '''Checks whether the value is a complex restriction (starts with the key character of a complex restriction)

    :param value: value to be checked
    :type value: string
    :return: boolean specifying the or_values list contains a complex restriction
    :rtype: boolean
    '''
    if value[:8] == 'pattern=' or value[:2] == '\\<' or value[:2] == '\\>' or value[:7] == 'length=' or value[:8] == 'length<=' or value[:8] == 'length>=':
        return True
    return False

def create_ids_specifications(ids_file, spec_list):
    '''Creates a new IDS specification for each entry in the spec_list and appends it to the ids_file.
    
    :param ids_file: the used ids_file
    :type ids_file: object ids
    :param spec_list: list containing all specifications as dictionaries
    :type spec_list: list
    :param ifc_version: list containing the ifc versions used for all specifications
    :type ifc_version: list
    :return: None
    '''
    i = 1
    for spec_data in spec_list:
        app_data = spec_data['app']
        string_instructions = ''
        if 'Phase' in spec_data['general']: string_instructions += 'Phase: ' + ', '.join(spec_data['general']['Phase']) + '; '
        if 'Role' in spec_data['general']: string_instructions += 'Role: ' + ', '.join(spec_data['general']['Role']) + '; '
        if 'Usecase' in spec_data['general']: string_instructions += 'Usecase: ' + ', '.join(spec_data['general']['Usecase']) + '; '
        string_instructions = string_instructions[0:len(string_instructions)-2]
        #Define specification name
        if STRING_SPECIFICATIONNAME in spec_data['spec']: spec_title = spec_data['spec'][STRING_SPECIFICATIONNAME][0]
        else:
            spec_title = 'Specification ' + str(i)
            if len(app_data) > 0:
                if STRING_ENTITY in app_data[0]:
                    spec_title += ':'
                    for j in range(len(app_data[0][STRING_ENTITY])):
                        spec_title += ' ' + app_data[0][STRING_ENTITY][j]
                        if STRING_PREDEFINEDTYPE in app_data[0]:
                            spec_title += '.' + app_data[0][STRING_PREDEFINEDTYPE][j]
        #Define specification cardinality
        spec_minOccurs=0
        spec_maxOccurs="unbounded"
        if STRING_SPECIFICATIONCARDINALITY in spec_data['spec']:
            if spec_data['spec'][STRING_SPECIFICATIONCARDINALITY][0].lower() == 'required':
                spec_minOccurs = 1
            if spec_data['spec'][STRING_SPECIFICATIONCARDINALITY][0].lower() == 'prohibited':
                spec_maxOccurs = 0
        #Define specification IFC version
        ifc_version = spec_data['spec'][STRING_SPECIFICATIONIFCVERSION]
        #Create specification
        ids_spec = ids.Specification(name=spec_title, ifcVersion=ifc_version, minOccurs=spec_minOccurs, maxOccurs=spec_maxOccurs, instructions=string_instructions if string_instructions != '' else None)
        #Append applicability and requirements
        append_facets(ids_spec.applicability, app_data)
        append_facets(ids_spec.requirements, spec_data['req'])
        #Append specification to ids file
        ids_file.specifications.append(ids_spec)
        i += 1

def append_facets(facets, input_data):
    '''Creates a new IDS facet for each entry in the input_data list and and stores it in a list (facets).
    It also converts enumerations, patterns, bounds, and lengths into ids resctrictions.

    :param facets: list to store the facets 
    :type facets: list
    :param input_data: list containing all facet data as dictionaries
    :type input_data: list
    :return: None
    '''

    #prepare the input data for using the ifcopenshell functions
    new_input_data = []
    for input_dict in input_data:
        new_input_dict = {}
        for key in input_dict.keys():
            or_values = input_dict[key]
            if isinstance(or_values, list):
                or_values = list(dict.fromkeys(or_values))
                restriction_base = 'string'
                #Determine the restriction base for complex restrictions depending on the datatype
                if key == STRING_PROPERTYVALUE and STRING_PROPERTYDATATYPE in input_dict:
                    restriction_base = datatype_base_dict[input_dict[STRING_PROPERTYDATATYPE][0]]
                #Complex restriction: enumeration  
                if len(or_values) > 1:
                    list_value = or_values[0]
                    if isinstance(list_value, str):
                        list_as_restriction = ids.Restriction(options={'enumeration': or_values}, base=restriction_base)
                        new_input_dict[key] = list_as_restriction
                elif len(or_values) == 1:
                    new_value = str(or_values[0])
                    #Boolean handling
                    if new_value.lower() == 'true' or new_value.lower() == 'false':
                        new_value = new_value.lower()
                    #Complex restriction: pattern
                    if new_value[:8] == 'pattern=':
                        new_value = ids.Restriction(options={'pattern': new_value[8:]}, base=restriction_base)
                    #Complex restriction: bounds
                    elif new_value[:3] == '\\<=':
                        if restriction_base == 'string': restriction_base = 'double'
                        new_value = ids.Restriction(options={'maxInclusive': float(new_value[3:].replace(',','.'))}, base=restriction_base)
                    elif new_value[:2] == '\\<':
                        if restriction_base == 'string': restriction_base = 'double'
                        new_value = ids.Restriction(options={'maxExclusive': float(new_value[2:].replace(',','.'))}, base=restriction_base)
                    elif new_value[:3] == '\\>=':
                        if restriction_base == 'string': restriction_base = 'double'
                        new_value = ids.Restriction(options={'minInclusive': float(new_value[3:].replace(',','.'))}, base=restriction_base)
                    elif new_value[:2] == '\\>':
                        if restriction_base == 'string': restriction_base = 'double'
                        new_value = ids.Restriction(options={'minExclusive': float(new_value[2:].replace(',','.'))}, base=restriction_base)
                    #Complex restriction: length
                    elif new_value[:7] == 'length=':
                        new_value = ids.Restriction(options={'length': int(float(new_value[7:].replace(',','.')))}, base=restriction_base)
                    elif new_value[:8] == 'length>=':
                        new_value = ids.Restriction(options={'minLength': int(float(new_value[8:].replace(',','.')))}, base=restriction_base)
                    elif new_value[:8] == 'length<=':
                        new_value = ids.Restriction(options={'maxLength': int(float(new_value[8:].replace(',','.')))}, base=restriction_base)
                    new_input_dict[key] = new_value
            else:
                #If or_values is None
                new_input_dict[key] = or_values
        new_input_data.append(new_input_dict)

    #Create an ids facet for each prepared input dictionary
    for input_dict in new_input_data:
        facet = None
        incomplete_facet = None
        #Entity facet
        if STRING_ENTITY in input_dict or STRING_PREDEFINEDTYPE in input_dict:
            if STRING_ENTITY not in input_dict: incomplete_facet = 'The Entity facet requires the Entity parameter.'
            else:
                key = str(input_dict[STRING_ENTITY])
                if STRING_PREDEFINEDTYPE in input_dict: key += '.' + str(input_dict[STRING_PREDEFINEDTYPE])
                facet = ids.Entity(name=input_dict[STRING_ENTITY] if STRING_ENTITY in input_dict else None,
                                predefinedType=input_dict[STRING_PREDEFINEDTYPE] if STRING_PREDEFINEDTYPE in input_dict else None,
                                instructions=input_dict[STRING_DESCRIPTION] if STRING_DESCRIPTION in input_dict else None)
        #Property facet
        elif STRING_PROPERTYSET in input_dict or STRING_PROPERTY in input_dict or STRING_PROPERTYDATATYPE in input_dict or STRING_PROPERTYVALUE in input_dict:
            if STRING_PROPERTYSET not in input_dict or STRING_PROPERTY not in input_dict: incomplete_facet = 'The Property facet requires the PropertySet and Property parameters.'
            else:
                key = str(input_dict[STRING_PROPERTY])
                facet = ids.Property(propertySet=input_dict[STRING_PROPERTYSET] if STRING_PROPERTYSET in input_dict else None,
                                    baseName=input_dict[STRING_PROPERTY] if STRING_PROPERTY in input_dict else None,
                                    dataType=input_dict[STRING_PROPERTYDATATYPE].replace(' ','') if STRING_PROPERTYDATATYPE in input_dict else None,
                                    value=input_dict[STRING_PROPERTYVALUE] if STRING_PROPERTYVALUE in input_dict else None,
                                    uri=input_dict[STRING_PROPERTYURI] if STRING_PROPERTYURI in input_dict else None,
                                    cardinality=input_dict[STRING_REQUIREMENTCARDINALITY] if STRING_REQUIREMENTCARDINALITY in input_dict else None,
                                    instructions=property_description_dict[key] if key in property_description_dict else None)
        #Material facet
        elif STRING_MATERIAL in input_dict:
            facet = ids.Material(value=input_dict[STRING_MATERIAL] if STRING_MATERIAL in input_dict else None,
                                uri=input_dict[STRING_MATERIALURI] if STRING_MATERIALURI in input_dict else None,
                                cardinality=input_dict[STRING_REQUIREMENTCARDINALITY] if STRING_REQUIREMENTCARDINALITY in input_dict else None)
        #Attribute facet
        elif STRING_ATTRIBUTE in input_dict or STRING_ATTRIBUTEVALUE in input_dict:
            if STRING_ATTRIBUTE not in input_dict: incomplete_facet = 'The Attribute facet requires the Attribute parameter.'
            else:
                facet = ids.Attribute(name=input_dict[STRING_ATTRIBUTE] if STRING_ATTRIBUTE in input_dict else None,
                                    value=input_dict[STRING_ATTRIBUTEVALUE] if STRING_ATTRIBUTEVALUE in input_dict else None,
                                    cardinality=input_dict[STRING_REQUIREMENTCARDINALITY] if STRING_REQUIREMENTCARDINALITY in input_dict else None)
        #Classification facet
        elif STRING_CLASSIFICATION in input_dict or STRING_CLASSIFICATIONSYSTEM in input_dict or STRING_CLASSIFICATIONURI in input_dict:
            if STRING_CLASSIFICATIONSYSTEM not in input_dict: incomplete_facet = 'The Classification facet requires the Classifiaction system parameter.'
            else:
                facet = ids.Classification(value=input_dict[STRING_CLASSIFICATION] if STRING_CLASSIFICATION in input_dict else None,
                                        system=input_dict[STRING_CLASSIFICATIONSYSTEM] if STRING_CLASSIFICATIONSYSTEM in input_dict else None,
                                        uri=input_dict[STRING_CLASSIFICATIONURI] if STRING_CLASSIFICATIONURI in input_dict else None,
                                        cardinality=input_dict[STRING_REQUIREMENTCARDINALITY] if STRING_REQUIREMENTCARDINALITY in input_dict else None)
        #PartOf facet
        elif STRING_PARTOFENTITY in input_dict or STRING_PARTOFPREDEFINEDTYPE in input_dict or STRING_PARTOFRELATION in input_dict:
            if STRING_PARTOFENTITY not in input_dict: incomplete_facet = 'The PartOf facet requires the Entity parameter.'
            else:
                facet = ids.PartOf(name=input_dict[STRING_PARTOFENTITY] if STRING_PARTOFENTITY in input_dict else None,
                                        predefinedType=input_dict[STRING_PARTOFPREDEFINEDTYPE] if STRING_PARTOFPREDEFINEDTYPE in input_dict else None,
                                        relation=input_dict[STRING_PARTOFRELATION] if STRING_PARTOFRELATION in input_dict else None,
                                        cardinality=input_dict[STRING_REQUIREMENTCARDINALITY] if STRING_REQUIREMENTCARDINALITY in input_dict else None)
        if incomplete_facet != None:
            raise Exception('Incomplete IDS facet: ' + str(input_dict) + '. ' + incomplete_facet)
        
        if facet != None:
            facets.append(facet)

def separate_specs_by_generaldata(input_data, separate_by_list):
    '''Separates the specifications in the input data list and assigns them to the different separators
    Returns a new dictionary with specification lists per separator

    :param input_data: List of specifications 
    :type input_data: list
    :param separate_by_list: List of all separators 
    :type separate_by_list: list
    :return: Dictionary with specification lists per separator
    :rtype: dict
    '''
    separated_data = {}
    for spec in input_data:
        #Initialize list of separation strings
        separation_strings = ['']
        #Create all possible combinations of general data separators in the current spec
        for key in spec['general'].keys():
            if key in separate_by_list:
                separation_strings_new = []
                for sep_string in separation_strings:
                    #for each separator in the general data of the current spec
                    for value in spec['general'][key]:
                        #append the separator to each prevoius sep_string
                        sep_string_new = sep_string + '_' + key + str(value)
                        separation_strings_new.append(sep_string_new)
                separation_strings= separation_strings_new
        
        #create new dictvalue or append spec to exisitng dict value for each separation string
        for sep_string in separation_strings:
            if sep_string in separated_data.keys():
                separated_data[sep_string]['specs'].append(spec)
                if 'Phase' in spec['general']: separated_data[sep_string]['general']['Phase'].update(spec['general']['Phase'])
                if 'Role' in spec['general']: separated_data[sep_string]['general']['Role'].update(spec['general']['Role'])
                if 'Usecase' in spec['general']: separated_data[sep_string]['general']['Usecase'].update(spec['general']['Usecase'])
            else:
                separated_data[sep_string] = {}
                separated_data[sep_string]['specs'] = [spec]
                separated_data[sep_string]['general'] = {}
                if 'Phase' in spec['general']: separated_data[sep_string]['general']['Phase'] = set(spec['general']['Phase'])
                else: separated_data[sep_string]['general']['Phase'] = set()
                if 'Role' in spec['general']: separated_data[sep_string]['general']['Role'] = set(spec['general']['Role'])
                else: separated_data[sep_string]['general']['Role'] = set()
                if 'Usecase' in spec['general']: separated_data[sep_string]['general']['Usecase'] = set(spec['general']['Usecase'])
                else: separated_data[sep_string]['general']['Usecase'] = set()
    return separated_data

def add_comment_to_xml(file_path, comment_text):
    # Parse the XML file
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    # Create the comment element
    comment = ET.Comment(comment_text)
    
    # Add the comment as the first child of the root
    root.insert(0, comment)
    
    # Write the modified XML with manual formatting
    with open(file_path, "wb") as f:
        # Add the XML declaration manually
        f.write(b'<?xml version="1.0" encoding="utf-8"?>\n')
        
        # Add the root element with the comment and newline
        f.write(ET.tostring(root, encoding="utf-8", method="xml").replace(b'-->', b'-->\n\t'))

merger = Merger(
    # Define custom merge strategies for different data types
    
    # First argument: List of merge strategies for specific types
    [(dict, 'merge'), (list, 'append_unique'), (set, 'union')],
    
    # Second argument: Fallback strategy for any other types (if they differ, replace)
    ['override'],
    
    # Third argument: List of strategies to apply to certain types in order of priority
    ['override']
)

datatype_base_dict = {
    'IFCABSORBEDDOSEMEASURE': 'double',
    'IFCACCELERATIONMEASURE': 'double',
    'IFCACTIONREQUESTTYPEENUM': 'string',
    'IFCACTIONSOURCETYPEENUM': 'string',
    'IFCACTIONTYPEENUM': 'string',
    'IFCACTUATORTYPEENUM': 'string',
    'IFCADDRESSTYPEENUM': 'string',
    'IFCAIRTERMINALBOXTYPEENUM': 'string',
    'IFCAIRTERMINALTYPEENUM': 'string',
    'IFCAIRTOAIRHEATRECOVERYTYPEENUM': 'string',
    'IFCALARMTYPEENUM': 'string',
    'IFCALIGNMENTCANTSEGMENTTYPEENUM': 'string',
    'IFCALIGNMENTHORIZONTALSEGMENTTYPEENUM': 'string',
    'IFCALIGNMENTTYPEENUM': 'string',
    'IFCALIGNMENTVERTICALSEGMENTTYPEENUM': 'string',
    'IFCAMOUNTOFSUBSTANCEMEASURE': 'double',
    'IFCANALYSISMODELTYPEENUM': 'string',
    'IFCANALYSISTHEORYTYPEENUM': 'string',
    'IFCANGULARVELOCITYMEASURE': 'double',
    'IFCANNOTATIONTYPEENUM': 'string',
    'IFCAREADENSITYMEASURE': 'double',
    'IFCAREAMEASURE': 'double',
    'IFCARITHMETICOPERATORENUM': 'string',
    'IFCASSEMBLYPLACEENUM': 'string',
    'IFCAUDIOVISUALAPPLIANCETYPEENUM': 'string',
    'IFCBEAMTYPEENUM': 'string',
    'IFCBEARINGTYPEENUM': 'string',
    'IFCBENCHMARKENUM': 'string',
    'IFCBINARY': '',
    'IFCBOILERTYPEENUM': 'string',
    'IFCBOOLEAN': 'boolean',
    'IFCBOXALIGNMENT': 'string',
    'IFCBRIDGEPARTTYPEENUM': 'string',
    'IFCBRIDGETYPEENUM': 'string',
    'IFCBUILDINGELEMENTPARTTYPEENUM': 'string',
    'IFCBUILDINGELEMENTPROXYTYPEENUM': 'string',
    'IFCBUILDINGSYSTEMTYPEENUM': 'string',
    'IFCBUILTSYSTEMTYPEENUM': 'string',
    'IFCBURNERTYPEENUM': 'string',
    'IFCCABLECARRIERFITTINGTYPEENUM': 'string',
    'IFCCABLECARRIERSEGMENTTYPEENUM': 'string',
    'IFCCABLEFITTINGTYPEENUM': 'string',
    'IFCCABLESEGMENTTYPEENUM': 'string',
    'IFCCAISSONFOUNDATIONTYPEENUM': 'string',
    'IFCCARDINALPOINTREFERENCE': 'integer',
    'IFCCHANGEACTIONENUM': 'string',
    'IFCCHILLERTYPEENUM': 'string',
    'IFCCHIMNEYTYPEENUM': 'string',
    'IFCCOILTYPEENUM': 'string',
    'IFCCOLUMNTYPEENUM': 'string',
    'IFCCOMMUNICATIONSAPPLIANCETYPEENUM': 'string',
    'IFCCOMPLEXPROPERTYTEMPLATETYPEENUM': 'string',
    'IFCCOMPRESSORTYPEENUM': 'string',
    'IFCCONDENSERTYPEENUM': 'string',
    'IFCCONNECTIONTYPEENUM': 'string',
    'IFCCONSTRAINTENUM': 'string',
    'IFCCONSTRUCTIONEQUIPMENTRESOURCETYPEENUM': 'string',
    'IFCCONSTRUCTIONMATERIALRESOURCETYPEENUM': 'string',
    'IFCCONSTRUCTIONPRODUCTRESOURCETYPEENUM': 'string',
    'IFCCONTEXTDEPENDENTMEASURE': 'double',
    'IFCCONTROLLERTYPEENUM': 'string',
    'IFCCONVEYORSEGMENTTYPEENUM': 'string',
    'IFCCOOLEDBEAMTYPEENUM': 'string',
    'IFCCOOLINGTOWERTYPEENUM': 'string',
    'IFCCOSTITEMTYPEENUM': 'string',
    'IFCCOSTSCHEDULETYPEENUM': 'string',
    'IFCCOUNTMEASURE': 'integer',
    'IFCCOURSETYPEENUM': 'string',
    'IFCCOVERINGTYPEENUM': 'string',
    'IFCCREWRESOURCETYPEENUM': 'string',
    'IFCCURRENCYENUM': 'string',
    'IFCCURTAINWALLTYPEENUM': 'string',
    'IFCCURVATUREMEASURE': 'double',
    'IFCCURVEINTERPOLATIONENUM': 'string',
    'IFCDAMPERTYPEENUM': 'string',
    'IFCDATAORIGINENUM': 'string',
    'IFCDATE': 'date',
    'IFCDATETIME': 'dateTime',
    'IFCDAYINMONTHNUMBER': 'integer',
    'IFCDAYINWEEKNUMBER': 'integer',
    'IFCDAYLIGHTSAVINGHOUR': 'integer',
    'IFCDERIVEDUNITENUM': 'string',
    'IFCDESCRIPTIVEMEASURE': 'string',
    'IFCDIMENSIONCOUNT': 'integer',
    'IFCDIRECTIONSENSEENUM': 'string',
    'IFCDISCRETEACCESSORYTYPEENUM': 'string',
    'IFCDISTRIBUTIONBOARDTYPEENUM': 'string',
    'IFCDISTRIBUTIONCHAMBERELEMENTTYPEENUM': 'string',
    'IFCDISTRIBUTIONPORTTYPEENUM': 'string',
    'IFCDISTRIBUTIONSYSTEMENUM': 'string',
    'IFCDOCUMENTCONFIDENTIALITYENUM': 'string',
    'IFCDOCUMENTSTATUSENUM': 'string',
    'IFCDOORPANELOPERATIONENUM': 'string',
    'IFCDOORPANELPOSITIONENUM': 'string',
    'IFCDOORSTYLECONSTRUCTIONENUM': 'string',
    'IFCDOORSTYLEOPERATIONENUM': 'string',
    'IFCDOORTYPEENUM': 'string',
    'IFCDOORTYPEOPERATIONENUM': 'string',
    'IFCDOSEEQUIVALENTMEASURE': 'double',
    'IFCDUCTFITTINGTYPEENUM': 'string',
    'IFCDUCTSEGMENTTYPEENUM': 'string',
    'IFCDUCTSILENCERTYPEENUM': 'string',
    'IFCDURATION': 'duration',
    'IFCDYNAMICVISCOSITYMEASURE': 'double',
    'IFCEARTHWORKSCUTTYPEENUM': 'string',
    'IFCEARTHWORKSFILLTYPEENUM': 'string',
    'IFCELECTRICAPPLIANCETYPEENUM': 'string',
    'IFCELECTRICCAPACITANCEMEASURE': 'double',
    'IFCELECTRICCHARGEMEASURE': 'double',
    'IFCELECTRICCONDUCTANCEMEASURE': 'double',
    'IFCELECTRICCURRENTENUM': 'string',
    'IFCELECTRICCURRENTMEASURE': 'double',
    'IFCELECTRICDISTRIBUTIONBOARDTYPEENUM': 'string',
    'IFCELECTRICDISTRIBUTIONPOINTFUNCTIONENUM': 'string',
    'IFCELECTRICFLOWSTORAGEDEVICETYPEENUM': 'string',
    'IFCELECTRICFLOWTREATMENTDEVICETYPEENUM': 'string',
    'IFCELECTRICGENERATORTYPEENUM': 'string',
    'IFCELECTRICHEATERTYPEENUM': 'string',
    'IFCELECTRICMOTORTYPEENUM': 'string',
    'IFCELECTRICRESISTANCEMEASURE': 'double',
    'IFCELECTRICTIMECONTROLTYPEENUM': 'string',
    'IFCELECTRICVOLTAGEMEASURE': 'double',
    'IFCELEMENTASSEMBLYTYPEENUM': 'string',
    'IFCELEMENTCOMPOSITIONENUM': 'string',
    'IFCENERGYMEASURE': 'double',
    'IFCENERGYSEQUENCEENUM': 'string',
    'IFCENGINETYPEENUM': 'string',
    'IFCENVIRONMENTALIMPACTCATEGORYENUM': 'string',
    'IFCEVAPORATIVECOOLERTYPEENUM': 'string',
    'IFCEVAPORATORTYPEENUM': 'string',
    'IFCEVENTTRIGGERTYPEENUM': 'string',
    'IFCEVENTTYPEENUM': 'string',
    'IFCEXTERNALSPATIALELEMENTTYPEENUM': 'string',
    'IFCFACILITYPARTCOMMONTYPEENUM': 'string',
    'IFCFACILITYUSAGEENUM': 'string',
    'IFCFANTYPEENUM': 'string',
    'IFCFASTENERTYPEENUM': 'string',
    'IFCFILTERTYPEENUM': 'string',
    'IFCFIRESUPPRESSIONTERMINALTYPEENUM': 'string',
    'IFCFLOWDIRECTIONENUM': 'string',
    'IFCFLOWINSTRUMENTTYPEENUM': 'string',
    'IFCFLOWMETERTYPEENUM': 'string',
    'IFCFONTSTYLE': 'string',
    'IFCFONTVARIANT': 'string',
    'IFCFONTWEIGHT': 'string',
    'IFCFOOTINGTYPEENUM': 'string',
    'IFCFORCEMEASURE': 'double',
    'IFCFREQUENCYMEASURE': 'double',
    'IFCFURNITURETYPEENUM': 'string',
    'IFCGASTERMINALTYPEENUM': 'string',
    'IFCGEOGRAPHICELEMENTTYPEENUM': 'string',
    'IFCGEOMETRICPROJECTIONENUM': 'string',
    'IFCGEOTECHNICALSTRATUMTYPEENUM': 'string',
    'IFCGLOBALLYUNIQUEID': 'string',
    'IFCGLOBALORLOCALENUM': 'string',
    'IFCGRIDTYPEENUM': 'string',
    'IFCHEATEXCHANGERTYPEENUM': 'string',
    'IFCHEATFLUXDENSITYMEASURE': 'double',
    'IFCHEATINGVALUEMEASURE': 'double',
    'IFCHOURINDAY': 'integer',
    'IFCHUMIDIFIERTYPEENUM': 'string',
    'IFCIDENTIFIER': 'string',
    'IFCILLUMINANCEMEASURE': 'double',
    'IFCIMPACTPROTECTIONDEVICETYPEENUM': 'string',
    'IFCINDUCTANCEMEASURE': 'double',
    'IFCINTEGER': 'integer',
    'IFCINTEGERCOUNTRATEMEASURE': 'integer',
    'IFCINTERCEPTORTYPEENUM': 'string',
    'IFCINTERNALOREXTERNALENUM': 'string',
    'IFCINVENTORYTYPEENUM': 'string',
    'IFCIONCONCENTRATIONMEASURE': 'double',
    'IFCISOTHERMALMOISTURECAPACITYMEASURE': 'double',
    'IFCJUNCTIONBOXTYPEENUM': 'string',
    'IFCKERBTYPEENUM': 'string',
    'IFCKINEMATICVISCOSITYMEASURE': 'double',
    'IFCLABEL': 'string',
    'IFCLABORRESOURCETYPEENUM': 'string',
    'IFCLAMPTYPEENUM': 'string',
    'IFCLANGUAGEID': 'string',
    'IFCLAYERSETDIRECTIONENUM': 'string',
    'IFCLENGTHMEASURE': 'double',
    'IFCLIGHTDISTRIBUTIONCURVEENUM': 'string',
    'IFCLIGHTEMISSIONSOURCEENUM': 'string',
    'IFCLIGHTFIXTURETYPEENUM': 'string',
    'IFCLINEARFORCEMEASURE': 'double',
    'IFCLINEARMOMENTMEASURE': 'double',
    'IFCLINEARSTIFFNESSMEASURE': 'double',
    'IFCLINEARVELOCITYMEASURE': 'double',
    'IFCLIQUIDTERMINALTYPEENUM': 'string',
    'IFCLOADGROUPTYPEENUM': 'string',
    'IFCLOGICAL': 'string',
    'IFCLOGICALOPERATORENUM': 'string',
    'IFCLUMINOUSFLUXMEASURE': 'double',
    'IFCLUMINOUSINTENSITYDISTRIBUTIONMEASURE': 'double',
    'IFCLUMINOUSINTENSITYMEASURE': 'double',
    'IFCMAGNETICFLUXDENSITYMEASURE': 'double',
    'IFCMAGNETICFLUXMEASURE': 'double',
    'IFCMARINEFACILITYTYPEENUM': 'string',
    'IFCMARINEPARTTYPEENUM': 'string',
    'IFCMASSDENSITYMEASURE': 'double',
    'IFCMASSFLOWRATEMEASURE': 'double',
    'IFCMASSMEASURE': 'double',
    'IFCMASSPERLENGTHMEASURE': 'double',
    'IFCMECHANICALFASTENERTYPEENUM': 'string',
    'IFCMEDICALDEVICETYPEENUM': 'string',
    'IFCMEMBERTYPEENUM': 'string',
    'IFCMINUTEINHOUR': 'integer',
    'IFCMOBILETELECOMMUNICATIONSAPPLIANCETYPEENUM': 'string',
    'IFCMODULUSOFELASTICITYMEASURE': 'double',
    'IFCMODULUSOFLINEARSUBGRADEREACTIONMEASURE': 'double',
    'IFCMODULUSOFROTATIONALSUBGRADEREACTIONMEASURE': 'double',
    'IFCMODULUSOFSUBGRADEREACTIONMEASURE': 'double',
    'IFCMOISTUREDIFFUSIVITYMEASURE': 'double',
    'IFCMOLECULARWEIGHTMEASURE': 'double',
    'IFCMOMENTOFINERTIAMEASURE': 'double',
    'IFCMONETARYMEASURE': 'double',
    'IFCMONTHINYEARNUMBER': 'integer',
    'IFCMOORINGDEVICETYPEENUM': 'string',
    'IFCMOTORCONNECTIONTYPEENUM': 'string',
    'IFCNAVIGATIONELEMENTTYPEENUM': 'string',
    'IFCNONNEGATIVELENGTHMEASURE': 'double',
    'IFCNORMALISEDRATIOMEASURE': 'double',
    'IFCNULLSTYLE': 'string',
    'IFCNULLSTYLEENUM': 'string',
    'IFCNUMERICMEASURE': 'double',
    'IFCOBJECTIVEENUM': 'string',
    'IFCOBJECTTYPEENUM': 'string',
    'IFCOCCUPANTTYPEENUM': 'string',
    'IFCOPENINGELEMENTTYPEENUM': 'string',
    'IFCOUTLETTYPEENUM': 'string',
    'IFCPARAMETERVALUE': 'double',
    'IFCPAVEMENTTYPEENUM': 'string',
    'IFCPERFORMANCEHISTORYTYPEENUM': 'string',
    'IFCPERMEABLECOVERINGOPERATIONENUM': 'string',
    'IFCPERMITTYPEENUM': 'string',
    'IFCPHMEASURE': 'double',
    'IFCPHYSICALORVIRTUALENUM': 'string',
    'IFCPILECONSTRUCTIONENUM': 'string',
    'IFCPILETYPEENUM': 'string',
    'IFCPIPEFITTINGTYPEENUM': 'string',
    'IFCPIPESEGMENTTYPEENUM': 'string',
    'IFCPLANARFORCEMEASURE': 'double',
    'IFCPLANEANGLEMEASURE': 'double',
    'IFCPLATETYPEENUM': 'string',
    'IFCPOSITIVEINTEGER': 'integer',
    'IFCPOSITIVELENGTHMEASURE': 'double',
    'IFCPOSITIVEPLANEANGLEMEASURE': 'double',
    'IFCPOSITIVERATIOMEASURE': 'double',
    'IFCPOWERMEASURE': 'double',
    'IFCPRESENTABLETEXT': 'string',
    'IFCPRESSUREMEASURE': 'double',
    'IFCPROCEDURETYPEENUM': 'string',
    'IFCPROFILETYPEENUM': 'string',
    'IFCPROJECTEDORTRUELENGTHENUM': 'string',
    'IFCPROJECTIONELEMENTTYPEENUM': 'string',
    'IFCPROJECTORDERRECORDTYPEENUM': 'string',
    'IFCPROJECTORDERTYPEENUM': 'string',
    'IFCPROPERTYSETTEMPLATETYPEENUM': 'string',
    'IFCPROPERTYSOURCEENUM': 'string',
    'IFCPROTECTIVEDEVICETRIPPINGUNITTYPEENUM': 'string',
    'IFCPROTECTIVEDEVICETYPEENUM': 'string',
    'IFCPUMPTYPEENUM': 'string',
    'IFCRADIOACTIVITYMEASURE': 'double',
    'IFCRAILINGTYPEENUM': 'string',
    'IFCRAILTYPEENUM': 'string',
    'IFCRAILWAYPARTTYPEENUM': 'string',
    'IFCRAILWAYTYPEENUM': 'string',
    'IFCRAMPFLIGHTTYPEENUM': 'string',
    'IFCRAMPTYPEENUM': 'string',
    'IFCRATIOMEASURE': 'double',
    'IFCREAL': 'double',
    'IFCRECURRENCETYPEENUM': 'string',
    'IFCREFERENTTYPEENUM': 'string',
    'IFCREFLECTANCEMETHODENUM': 'string',
    'IFCREINFORCEDSOILTYPEENUM': 'string',
    'IFCREINFORCINGBARROLEENUM': 'string',
    'IFCREINFORCINGBARSURFACEENUM': 'string',
    'IFCREINFORCINGBARTYPEENUM': 'string',
    'IFCREINFORCINGMESHTYPEENUM': 'string',
    'IFCRESOURCECONSUMPTIONENUM': 'string',
    'IFCRIBPLATEDIRECTIONENUM': 'string',
    'IFCROADPARTTYPEENUM': 'string',
    'IFCROADTYPEENUM': 'string',
    'IFCROLEENUM': 'string',
    'IFCROOFTYPEENUM': 'string',
    'IFCROTATIONALFREQUENCYMEASURE': 'double',
    'IFCROTATIONALMASSMEASURE': 'double',
    'IFCROTATIONALSTIFFNESSMEASURE': 'double',
    'IFCSANITARYTERMINALTYPEENUM': 'string',
    'IFCSECONDINMINUTE': 'double',
    'IFCSECTIONALAREAINTEGRALMEASURE': 'double',
    'IFCSECTIONMODULUSMEASURE': 'double',
    'IFCSECTIONTYPEENUM': 'string',
    'IFCSENSORTYPEENUM': 'string',
    'IFCSEQUENCEENUM': 'string',
    'IFCSERVICELIFEFACTORTYPEENUM': 'string',
    'IFCSERVICELIFETYPEENUM': 'string',
    'IFCSHADINGDEVICETYPEENUM': 'string',
    'IFCSHEARMODULUSMEASURE': 'double',
    'IFCSIGNALTYPEENUM': 'string',
    'IFCSIGNTYPEENUM': 'string',
    'IFCSIMPLEPROPERTYTEMPLATETYPEENUM': 'string',
    'IFCSLABTYPEENUM': 'string',
    'IFCSOLARDEVICETYPEENUM': 'string',
    'IFCSOLIDANGLEMEASURE': 'double',
    'IFCSOUNDPOWERLEVELMEASURE': 'double',
    'IFCSOUNDPOWERMEASURE': 'double',
    'IFCSOUNDPRESSURELEVELMEASURE': 'double',
    'IFCSOUNDPRESSUREMEASURE': 'double',
    'IFCSOUNDSCALEENUM': 'string',
    'IFCSPACEHEATERTYPEENUM': 'string',
    'IFCSPACETYPEENUM': 'string',
    'IFCSPATIALZONETYPEENUM': 'string',
    'IFCSPECIFICHEATCAPACITYMEASURE': 'double',
    'IFCSPECULAREXPONENT': 'double',
    'IFCSPECULARROUGHNESS': 'double',
    'IFCSTACKTERMINALTYPEENUM': 'string',
    'IFCSTAIRFLIGHTTYPEENUM': 'string',
    'IFCSTAIRTYPEENUM': 'string',
    'IFCSTATEENUM': 'string',
    'IFCSTRIPPEDOPTIONAL': 'boolean',
    'IFCSTRUCTURALCURVEACTIVITYTYPEENUM': 'string',
    'IFCSTRUCTURALCURVEMEMBERTYPEENUM': 'string',
    'IFCSTRUCTURALCURVETYPEENUM': 'string',
    'IFCSTRUCTURALSURFACEACTIVITYTYPEENUM': 'string',
    'IFCSTRUCTURALSURFACEMEMBERTYPEENUM': 'string',
    'IFCSTRUCTURALSURFACETYPEENUM': 'string',
    'IFCSUBCONTRACTRESOURCETYPEENUM': 'string',
    'IFCSURFACEFEATURETYPEENUM': 'string',
    'IFCSURFACETEXTUREENUM': 'string',
    'IFCSWITCHINGDEVICETYPEENUM': 'string',
    'IFCSYSTEMFURNITUREELEMENTTYPEENUM': 'string',
    'IFCTANKTYPEENUM': 'string',
    'IFCTASKDURATIONENUM': 'string',
    'IFCTASKTYPEENUM': 'string',
    'IFCTEMPERATUREGRADIENTMEASURE': 'double',
    'IFCTEMPERATURERATEOFCHANGEMEASURE': 'double',
    'IFCTENDONANCHORTYPEENUM': 'string',
    'IFCTENDONCONDUITTYPEENUM': 'string',
    'IFCTENDONTYPEENUM': 'string',
    'IFCTEXT': 'string',
    'IFCTEXTALIGNMENT': 'string',
    'IFCTEXTDECORATION': 'string',
    'IFCTEXTFONTNAME': 'string',
    'IFCTEXTTRANSFORMATION': 'string',
    'IFCTHERMALADMITTANCEMEASURE': 'double',
    'IFCTHERMALCONDUCTIVITYMEASURE': 'double',
    'IFCTHERMALEXPANSIONCOEFFICIENTMEASURE': 'double',
    'IFCTHERMALLOADSOURCEENUM': 'string',
    'IFCTHERMALLOADTYPEENUM': 'string',
    'IFCTHERMALRESISTANCEMEASURE': 'double',
    'IFCTHERMALTRANSMITTANCEMEASURE': 'double',
    'IFCTHERMODYNAMICTEMPERATUREMEASURE': 'double',
    'IFCTIME': 'time',
    'IFCTIMEMEASURE': 'double',
    'IFCTIMESERIESDATATYPEENUM': 'string',
    'IFCTIMESERIESSCHEDULETYPEENUM': 'string',
    'IFCTIMESTAMP': 'integer',
    'IFCTORQUEMEASURE': 'double',
    'IFCTRACKELEMENTTYPEENUM': 'string',
    'IFCTRANSFORMERTYPEENUM': 'string',
    'IFCTRANSPORTELEMENTTYPEENUM': 'string',
    'IFCTUBEBUNDLETYPEENUM': 'string',
    'IFCUNITARYCONTROLELEMENTTYPEENUM': 'string',
    'IFCUNITARYEQUIPMENTTYPEENUM': 'string',
    'IFCUNITENUM': 'string',
    'IFCURIREFERENCE': 'string',
    'IFCVALVETYPEENUM': 'string',
    'IFCVAPORPERMEABILITYMEASURE': 'double',
    'IFCVEHICLETYPEENUM': 'string',
    'IFCVIBRATIONDAMPERTYPEENUM': 'string',
    'IFCVIBRATIONISOLATORTYPEENUM': 'string',
    'IFCVIRTUALELEMENTTYPEENUM': 'string',
    'IFCVOIDINGFEATURETYPEENUM': 'string',
    'IFCVOLUMEMEASURE': 'double',
    'IFCVOLUMETRICFLOWRATEMEASURE': 'double',
    'IFCWALLTYPEENUM': 'string',
    'IFCWARPINGCONSTANTMEASURE': 'double',
    'IFCWARPINGMOMENTMEASURE': 'double',
    'IFCWASTETERMINALTYPEENUM': 'string',
    'IFCWELLKNOWNTEXTLITERAL': 'string',
    'IFCWINDOWPANELOPERATIONENUM': 'string',
    'IFCWINDOWPANELPOSITIONENUM': 'string',
    'IFCWINDOWSTYLECONSTRUCTIONENUM': 'string',
    'IFCWINDOWSTYLEOPERATIONENUM': 'string',
    'IFCWINDOWTYPEENUM': 'string',
    'IFCWINDOWTYPEPARTITIONINGENUM': 'string',
    'IFCWORKCALENDARTYPEENUM': 'string',
    'IFCWORKCONTROLTYPEENUM': 'string',
    'IFCWORKPLANTYPEENUM': 'string',
    'IFCWORKSCHEDULETYPEENUM': 'string',
    'IFCYEARNUMBER': 'integer',
}