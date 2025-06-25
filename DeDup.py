import pandas as pd
import sys
import numpy as np
import xml.etree.ElementTree as ET
import re
from fuzzywuzzy import fuzz, process
import math
from tqdm import tqdm
import time
import argparse
import os
import shutil


def validate_file(file_path):
    if not (file_path.endswith('.xlsx') or file_path.endswith('.csv')):
        raise argparse.ArgumentTypeError(f"File '{file_path}' must be a .xlsx or .csv file.")
    if not os.path.exists(file_path):
        raise argparse.ArgumentTypeError(f"File '{file_path}' does not exist.")
    if not os.access(file_path, os.R_OK):
        raise argparse.ArgumentTypeError(f"File '{file_path}' is not accessible (read permissions required).")
    return file_path

def validate_job_name(job_name):
    # Basic check: must be a valid folder name (no special characters typically restricted in folders)
    if re.search(r'[<>:"/\\|?*]', job_name):
        raise argparse.ArgumentTypeError(f"Job name '{job_name}' contains invalid characters for a folder name.")
    if os.path.exists(job_name):
        raise argparse.ArgumentTypeError(f"Job name '{job_name}' already exists as a folder. Please choose another name.")
    return job_name
    
    
def extract_variables_from_file(file_path="config.txt"):
    with open(file_path, "r", encoding="utf-8") as file:
        xml_content = file.read().strip()

    # Parse XML content
    root = ET.fromstring(xml_content)

    # Dictionary to store extracted data
    conditions = {}

    # Regex to match element names that start with "C" (e.g., C1, C2, C2a)
    #pattern = re.compile(r"^C\d+[a-zA-Z]*$")
    pattern = re.compile(r"^C\d+[a-zA-Z0-9]*$")


    # Iterate over all elements inside BASE_CONDITION
    for element in root:
        if pattern.match(element.tag):  # Check if element name starts with "C"
            text_content = element.text.strip() if element.text else ""

            # Extract variable-value pairs
            var_values = []
            for pair in text_content.split():
                if ":" in pair:
                    var, val = pair.split(":")
                    var_values.append((var.strip(), val.strip()))

            conditions[element.tag] = var_values

    return conditions
    
def compare_excel_columns(file1, file2):
    """
    Compares column names between two Excel files and prints an error if they do not match.
    Exits the program if there is a mismatch.
    
    Parameters:
    - file1 (str): Path to the first Excel file.
    - file2 (str): Path to the second Excel file.
    """
    df1 = pd.read_excel(file1, sheet_name=0)
    df2 = pd.read_excel(file2, sheet_name=0)
    
    columns1 = set(df1.columns)
    columns2 = set(df2.columns)
    
    missing_in_file1 = columns2 - columns1
    missing_in_file2 = columns1 - columns2
    
    if missing_in_file1 or missing_in_file2:
        if missing_in_file1:
            print(f"Error: The following columns are missing in {file1} but present in {file2}: {missing_in_file1}")
        if missing_in_file2:
            print(f"Error: The following columns are missing in {file2} but present in {file1}: {missing_in_file2}")
        sys.exit(1)
    
    print("Success: Both files contain the same column names.")

    
def create_hashmap_from_dataframe(df, key_column, value_column):
    """
    Creates a hashmap (dictionary) from a DataFrame.

    Parameters:
    - df (pd.DataFrame): Input DataFrame.
    - key_column (str): Column to be used as keys.
    - value_column (str): Column to be used as values.

    Returns:
    - dict: Hashmap created from the DataFrame.
    """
    if key_column not in df.columns or value_column not in df.columns:
        raise ValueError("Specified key or value column not found in DataFrame")

    return dict(zip(df[key_column], df[value_column]))
    
def load_excel_as_dataframe(file_path, sheet_name=0):
    """
    Loads an Excel (.xlsx) file into a Pandas DataFrame.

    Parameters:
    - file_path (str): Path to the Excel file.
    - sheet_name (str or int, optional): Sheet name or index to load (default is the first sheet).

    Returns:
    - pd.DataFrame: DataFrame containing the Excel sheet data.
    """
    return pd.read_excel(file_path, sheet_name=sheet_name)
    
    
def validate_excel_columns(rules_file, excel_file, output_file, warn):
    """
    Validates the existence of required columns and their data types in an Excel file.
    
    Parameters:
    - rules_file (str): Path to the CSV or Excel file containing validation rules.
    - excel_file (str): Path to the Excel file to be checked.
    - output_file (str): Path to save the validated Excel file if no errors are found.
    - warn (str): Warning message prefix for all error messages.
    
    Exits with an error message if any validation fails; otherwise, saves the validated data.
    """

    # Load validation rules from CSV or Excel
    if rules_file.endswith(".csv"):
        df_rules = pd.read_csv(rules_file)
    elif rules_file.endswith(".xlsx"):
        df_rules = pd.read_excel(rules_file, sheet_name=0)
    else:
        print(f"{warn}: Rules file must be in CSV or Excel format.")
        sys.exit(1)


    
    # Check required columns in the rules file
    required_columns = {"VARIABLE", "DATA_TYPE", "EQU_C_NAME"}
    if not required_columns.issubset(df_rules.columns):
        print(f"{warn}: Rules file must contain 'VARIABLE', 'DATA_TYPE', and 'EQU_C_NAME' columns.")
        sys.exit(1)

    # Load Excel file (assuming first sheet)
    df_excel = pd.read_excel(excel_file, sheet_name=0, dtype=str)  # Read as strings to validate types

    # Replace multiple consecutive spaces with a single space in all columns
    df_excel = df_excel.applymap(lambda x: ' '.join(str(x).split()) if pd.notna(x) else x)

    # Identify special columns based on VARIABLE values
    special_columns = set(df_rules.loc[df_rules["VARIABLE"].isin(["FULL_NAME", "FULL_ADDRESS", "RELATIVE"]), "EQU_C_NAME"])
    
    # Convert NaN values in special columns to string
    for col in special_columns.intersection(df_excel.columns):
        df_excel[col] = df_excel[col].fillna("" if col in special_columns else np.nan).astype(str)

    # Find columns corresponding to mandatory fields in EQU_C_NAME
    mandatory_variables = ["FULL_NAME", "AGE", "ICD", "GENDER", "REF"]
    mandatory_columns = df_rules.loc[df_rules["VARIABLE"].isin(mandatory_variables), "EQU_C_NAME"].values
    
    for col in mandatory_columns:
        if col in df_excel.columns:
            empty_rows = df_excel[df_excel[col].isna() | (df_excel[col].str.strip() == "")]
            if not empty_rows.empty:
                print(f"{warn}: '{col}' cannot be empty. The following rows have empty values:")
                print(empty_rows)
                sys.exit(1)

    # Validate column existence in Excel file
    missing_equ_c_names = [col for col in df_rules["EQU_C_NAME"] if col not in df_excel.columns]

    if missing_equ_c_names:
        print(f"{warn}: The following columns from 'EQU_C_NAME' are missing in the Excel file: {missing_equ_c_names}")
        sys.exit(1)

    # Data type validation
    for _, row in df_rules.iterrows():
        col_name = row["EQU_C_NAME"]
        expected_type = row["DATA_TYPE"]

        if col_name not in df_excel.columns:
            continue

        if expected_type == "UAN":  # Unique Alphanumeric
            if not df_excel[col_name].astype(str).is_unique:
                print(f"{warn}: Column '{col_name}' must have unique values.")
                print("Invalid duplicate values:")
                print(df_excel[df_excel.duplicated(col_name, keep=False)])
                sys.exit(1)
            invalid_values = df_excel[~df_excel[col_name].apply(lambda x: isinstance(x, str))]
            if not invalid_values.empty:
                print(f"{warn}: Column '{col_name}' must contain alphanumeric values.")
                print("Invalid values found:")
                print(invalid_values)
                sys.exit(1)

        elif expected_type == "N":  # Numeric
            df_excel[col_name] = pd.to_numeric(df_excel[col_name], errors='coerce')
            invalid_values = df_excel[df_excel[col_name].isna() & df_excel[col_name].notnull()]
            if not invalid_values.empty:
                print(f"{warn}: Column '{col_name}' must contain only numeric values.")
                print("Invalid values found:")
                print(invalid_values)
                sys.exit(1)

    # Save validated data to a new Excel file
    df_excel.to_excel(output_file, index=False)
    print(f"{warn}: All verifications OK. Validated data saved to {output_file}")
    # Create dictionary from VARIABLE and EQU_C_NAME columns
    variable_equ_cname_map = dict(zip(df_rules["VARIABLE"], df_rules["EQU_C_NAME"]))
    return variable_equ_cname_map

######################## MATCH FUCNTIONS ################################


def compare_ICD10(hashmap1, hashmap2, key1, key2):

    def read_file_as_list(file_path):
        """
        Reads a text file and returns its content as a list, where each line is an item.
        If the file is empty, returns ["XXX"].

        Parameters:
        - file_path (str): Path to the text file.

        Returns:
        - list: List of lines from the file or ["XXX"] if the file is empty.
        """
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                lines = file.read().splitlines()  # Read lines without '\n'
                return lines if lines else ["XXX"]  # Return content or ["XXX"] if empty
        except FileNotFoundError:
            print(f"Error: The file '{file_path}' was not found.")
            sys.exit(1)

    def compare_string_with_list(string, string_list):
        if string in string_list:
            return 5
        else:
            return -15
            
            
    value1 = hashmap1.get(key1)
    value2 = hashmap2.get(key2)
    value1= value1[:3]
    value2 = value2[:3]

    Unknown_ICD10 = read_file_as_list("ICD.txt") #['C26', 'C39','C48', 'C75', 'C76', 'C77', 'C78', 'C79', 'C97','C80']

    if value1 is not None and value2 is not None:
        if value1 in Unknown_ICD10:
            score = 5
        else:
            score = fuzz.token_sort_ratio(value1, value2)
            if score == 100:
                score = 5
            else:
                score = compare_string_with_list(value2, Unknown_ICD10)
    else:
        score = 0
    return score

def contains_characters(input_string):
    # Check if input_string is None, NaN, empty, or contains only whitespace
    if input_string is None or (isinstance(input_string, float) and math.isnan(input_string)) or input_string.strip() == '':
        return False
    else:
        return True

def fuzzywuzzy_compare(hashmap1, hashmap2, key1, key2):
    value1 = hashmap1.get(key1)
    value2 = hashmap2.get(key2)
    
    if contains_characters(value1) or contains_characters(value2) == True:
        if value1 is not None and value2 is not None:
            score = fuzz.token_sort_ratio(value1, value2)
        else:
            score = 0
    else:
        score = 0

    return score

def compare_relative(hashmap1, hashmap2, key1, key2):
    value1 = hashmap1.get(key1)
    value2 = hashmap2.get(key2)

    # Check if value1 or value2 is None, NaN, empty, or contains only whitespace
    if (value1 is None or (isinstance(value1, float) and math.isnan(value1)) or str(value1).strip() == '') or (value2 is None or (isinstance(value2, float) and math.isnan(value2)) or str(value2).strip() == ''):
        return 0

    return fuzz.token_sort_ratio(value1, value2)
    

# Function to check if a value is None, NaN, empty, or contains only whitespace
def is_invalid(value):
    return value is None or (isinstance(value, float) and math.isnan(value)) or str(value).strip() == ''

def compare_age(hashmap1, hashmap2, key1, key2):
    value1 = hashmap1.get(key1)
    value2 = hashmap2.get(key2)


    if is_invalid(value1) or is_invalid(value2):
        return -15

    try:
        score = math.sqrt(pow((int(value1) - int(value2)), 2))
        if score == 0:
            return 10
        elif score > 5:
            return -15
        else:
            return 0
    except ValueError:
        return -15  # Handles cases where conversion to int fails    
        
        
def compare_sex(hashmap1, hashmap2, key1, key2):
    value1 = hashmap1.get(key1)
    value2 = hashmap2.get(key2)

    if is_invalid(value1) or is_invalid(value2):
        return -10

    score = fuzz.token_sort_ratio(value1, value2)
    return 0 if score == 100 else -10

def compare_pin(hashmap1, hashmap2, key1, key2):
    value1 = hashmap1.get(key1)
    value2 = hashmap2.get(key2)

    if is_invalid(value1) or is_invalid(value2):
        return -10

    score = fuzz.token_sort_ratio(str(value1), str(value2))
    return 0 if score == 100 else -10    

def save_dataframe_to_excel(df, filename):
    """
    Saves a Pandas DataFrame to an Excel (.xlsx) file.

    Parameters:
    - df (pd.DataFrame): The DataFrame to save.
    - filename (str): Name of the output file (default: 'output.xlsx').

    Returns:
    - None
    """
    try:
        df.to_excel(filename, index=False, engine='openpyxl')  # Avoids writing row indices
        print(f"DataFrame successfully saved as '{filename}'")
    except Exception as e:
        print(f"Error saving DataFrame: {e}")


def extract_matching_records(file1, file2, output_file):
    # Load the two Excel files
    df1 = pd.read_excel(file1)  # File 1 containing Q_ID, T_ID, FULL_NAME, etc.
    df2 = pd.read_excel(file2)  # File 2 containing REF, FULL_NAME, etc.

    # Ensure required columns exist
    required_columns1 = {"Q_ID", "T_ID", "TOTAL", "FULL_NAME", "FULL_ADDRESS", "AGE", "GENDER", "ICD", "PINCODE", "RELATIVE"}
    required_columns2 = {"REF", "FULL_NAME", "FULL_ADDRESS", "RELATIVE", "ICD", "AGE", "GENDER", "PINCODE"}

    if not required_columns1.issubset(df1.columns) or not required_columns2.issubset(df2.columns):
        raise ValueError("Required columns are missing in one of the input files.")

    result_df = pd.DataFrame()  # Empty DataFrame to store results

    for _, row in df1.iterrows():
        q_id, t_id = row["Q_ID"], row["T_ID"]

        # Extract matching rows from File 2 where REF matches Q_ID
        qid_matches = df2[df2["REF"] == q_id]
        # Extract matching rows from File 2 where REF matches T_ID
        tid_matches = df2[df2["REF"] == t_id]

        if not qid_matches.empty or not tid_matches.empty:
            # Create a blank row in File 2 format
            blank_row = pd.DataFrame([[""] * len(df2.columns)], columns=df2.columns)

            # Fill the blank row with TOTAL from File 1
            blank_row["REF"] = f"{row['TOTAL']}({row['PROB']})"

            # Copy matching column values from File 1 to File 2 format
            for col in df1.columns:
                if col in df2.columns:  # Only insert if column exists in File 2
                    blank_row[col] = row[col]

            # Append extracted data, then the blank row
            result_df = pd.concat([
                result_df, 
                qid_matches,  # Append Q_ID matches
                tid_matches,  # Append T_ID matches
                blank_row     # Append blank row with TOTAL & matched columns
            ], ignore_index=True)

    # Save the extracted records to an Excel file
    result_df.to_excel(output_file, index=False)
    print(f"Extracted data saved to {output_file}")

def score_merger(input_file, output_file, col1, col2):
    # Read the input Excel file
    print("\nMerging Score....")
    df = pd.read_excel(input_file)

    # Merge the values from col1 and col2 into a new column
    df['merged'] = df[col1].astype(str) + " " + df[col2].astype(str)

    # Save the modified DataFrame to an Excel file
    df.to_excel(output_file, index=False)

def score_remover(input_file, output_file, column):
    # Read the input Excel file
    df = pd.read_excel(input_file)
    iden = []
    n = 0
    Fn = len(df)
    print("\nRemoving Identical pairs....")

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        n += 1
        current_value = row[column]
        for i in range(index + 1, len(df)):
            other_value = df.loc[i, column]
            match_score = fuzz.token_sort_ratio(current_value, other_value)
            if match_score == 100:
                iden.append(i)
                break
        

    # Drop duplicate rows and remove the specified column
    df = df.drop(iden)
    df.drop(columns=[column], inplace=True)

    # Save the modified DataFrame to an Excel file
    df.to_excel(output_file, index=False)   

def sort_excel_descending(input_file, output_file, column_name):
    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Sort the DataFrame in descending order of the specified column
    df = df.sort_values(by=column_name, ascending=False)

    # Save the sorted DataFrame to an Excel file
    df.to_excel(output_file, index=False)    

def move_files_to_folder(folder_name, file_list):
    # Create the folder
    # try:
        # os.mkdir(folder_name)
    # except OSError as e:
        # print(f"Error creating folder: {e}")
        # return

    # Check write permission for the folder
    if not os.access(folder_name, os.W_OK):
        print(f"No write permission for folder: {folder_name}")
        return

    # Move the files to the folder
    for file_path in file_list:
        try:
            shutil.move(file_path, folder_name)
        except Exception as e:
            print(f"Error moving file '{file_path}' to folder: {e}")

    print("File move completed successfully.")       
############################################


### Main Begins ###############



def main():
#Validate column names
    parser = argparse.ArgumentParser(description="Validate two input files and a job name.")
    parser.add_argument('-f1', required=True, type=validate_file, help='First input file (.xlsx or .csv)')
    parser.add_argument('-f2', required=True, type=validate_file, help='Second input file (.xlsx or .csv)')
    parser.add_argument('-j', required=True, type=validate_job_name, help='Job name (valid, non-existing folder name)')

    args = parser.parse_args()

    print(f"File 1 is valid: {args.f1}")
    print(f"File 2 is valid: {args.f2}")
    print(f"Job name is valid and available: {args.j}")
    os.mkdir(args.j)

    Q_col = validate_excel_columns("S_column.xlsx", args.f1,"QC.xlsx","Query_File")
    T_col = validate_excel_columns("S_column.xlsx", args.f2,"TC.xlsx","Target_File")
    compare_excel_columns("QC.xlsx","TC.xlsx")

#Create dataframe for the cleaned .xlsx
    Q_df = load_excel_as_dataframe("QC.xlsx")
    T_df = load_excel_as_dataframe("TC.xlsx")

#Create hash for comparison
    Q_hash_NAME = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['FULL_NAME'])
    T_hash_NAME = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['FULL_NAME'])
    Q_hash_ADDRESS = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['FULL_ADDRESS'])
    T_hash_ADDRESS = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['FULL_ADDRESS'])
    Q_hash_AGE = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['AGE'])
    T_hash_AGE = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['AGE'])
    Q_hash_GENDER = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['GENDER'])
    T_hash_GENDER = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['GENDER'])
    Q_hash_ICD = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['ICD'])
    T_hash_ICD = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['ICD'])
    Q_hash_PIN = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['PINCODE'])
    T_hash_PIN = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['PINCODE'])
    Q_hash_RELATIVE = create_hashmap_from_dataframe(Q_df,Q_col['REF'],Q_col['RELATIVE'])
    T_hash_RELATIVE = create_hashmap_from_dataframe(T_df,T_col['REF'],T_col['RELATIVE'])


#debug_xml_structure()
    config_data = extract_variables_from_file("config.txt")
    #print(config_data)
    #input()
    compared_rows = pd.DataFrame(columns=['Q_ID','T_ID', Q_col['FULL_NAME'],Q_col['FULL_ADDRESS'],Q_col['AGE'],Q_col['GENDER'],Q_col['ICD'],Q_col['PINCODE'],Q_col['RELATIVE'], "TOTAL" ,"PROB"])


#Progress Bar
    Query_iterations = len(Q_hash_NAME)
    Target_iterations = len(T_hash_NAME)
    total_steps = Query_iterations * Target_iterations  
    with tqdm(total=total_steps, desc="Comparing", unit="step") as pbar:
        for key1, value1 in Q_hash_NAME.items():
                for key2, value2 in T_hash_NAME.items():
                    if(key1 != key2):
                        C1=C1a=C3b2=C3b8=C4d=C4e2 = fuzz.token_sort_ratio(value1, value2) #FULL_NAME MATCH
                        if C1 >=50:
                            C1b=C3b6=C4c = compare_ICD10(Q_hash_ICD,T_hash_ICD,key1,key2) #ICD MATCH
                            C2=C5=C3b9 = fuzzywuzzy_compare(Q_hash_ADDRESS,T_hash_ADDRESS,key1,key2) # FULL_ADDRESS MATCH
                            C3b1=C5a=C4e= compare_relative(Q_hash_RELATIVE,T_hash_RELATIVE,key1,key2) # RELATIVE MATCH
                            C3b4=C4a=C5c = compare_age(Q_hash_AGE,T_hash_AGE,key1,key2) # AGE MATCH
                            C3b5=C4b=C5b = compare_sex(Q_hash_GENDER,T_hash_GENDER,key1,key2) # GENDER MATCH
                            C3b7=C4e1=C5d = compare_pin(Q_hash_PIN,T_hash_PIN,key1,key2) # PIN MATCH
                            C3a=C3b=C3b3=C5e=Total_S= C1+C1b+C2+C3b4+C3b5+C3b7 # Total Match Score
                            P1=P2=P3=P4=P5=P6=P7=0
                            
                            P1 = 1 if C1 >= 75 else 0 #Name
                            P2 = 1 if C2 >= 70 else 0 #Add
                            if(C3b4 > 9):
                                P3 = 1
                            elif(C3b4 < 0 ):
                                P3 = 0.5
                            else:
                                P3 = 0
                            P4 = 1 if C1b >0 else 0 #ICD
                            P5 = 1 if C3b7 >= 0 else 0 #PIN
                            P6 = 1 if C3b5 >= 0 else 0 #GEN
                            P7 = 1 if Total_S >=145 else 0
                            Norm_score = (P1+P2+P3+P4+P5+P6+P7)/7 #Proability Score
                            if (Norm_score >= float(config_data.get('C0',[])[0][1])):
############# Nested conditions begin ##############################
                                if(C1 >= int(config_data.get('C1', [])[0][1])) or (C1b >= int(config_data.get('C1b', [])[0][1]) and (C1a >= int(config_data.get('C1a', [])[0][1]))):
                                    if(C2 >= int(config_data.get('C2', [])[0][1])):
                                        if C3a >=  int(config_data.get('C3a', [])[0][1]):
                                            compared_rows = compared_rows._append({
                                                'Q_ID': key1,
                                                'T_ID': key2,
                                                Q_col['FULL_NAME']: C1,
                                                Q_col['FULL_ADDRESS']: C2,
                                                Q_col['AGE']: C3b4,
                                                Q_col['GENDER']: C3b5,
                                                Q_col['PINCODE']: C3b7,
                                                Q_col['ICD']: C1b,
                                                Q_col['RELATIVE']: C3b1,
                                                'TOTAL': Total_S,
                                                'PROB': Norm_score
                                            }, ignore_index=True)
                                        elif C3b >= int(config_data.get('C3b', [])[0][1]):
                                            if C3b1 >= int(config_data.get('C3b1', [])[0][1]):
                                                compared_rows = compared_rows._append({
                                                    'Q_ID': key1,
                                                    'T_ID': key2,
                                                    Q_col['FULL_NAME']: C1,
                                                    Q_col['FULL_ADDRESS']: C2,
                                                    Q_col['AGE']: C3b4,
                                                    Q_col['GENDER']: C3b5,
                                                    Q_col['PINCODE']: C3b7,
                                                    Q_col['ICD']: C1b,
                                                    Q_col['RELATIVE']: C3b1,
                                                    'TOTAL': Total_S,
                                                    'PROB': Norm_score
                                                }, ignore_index=True)
                                            elif C3b2 >= int(config_data.get('C3b2', [])[0][1]) and C3b3 >= int(config_data.get('C3b3', [])[0][1]):
                                                compared_rows = compared_rows._append({
                                                    'Q_ID': key1,
                                                    'T_ID': key2,
                                                    Q_col['FULL_NAME']: C1,
                                                    Q_col['FULL_ADDRESS']: C2,
                                                    Q_col['AGE']: C3b4,
                                                    Q_col['GENDER']: C3b5,
                                                    Q_col['PINCODE']: C3b7,
                                                    Q_col['ICD']: C1b,
                                                    Q_col['RELATIVE']: C3b1,
                                                    'TOTAL': Total_S,
                                                    'PROB': Norm_score
                                                }, ignore_index=True)                            
                                            elif  C3b4 >= int(config_data.get('C3b4', [])[0][1]) and C3b5 >= int(config_data.get('C3b5', [])[0][1]) and C3b6 >= int(config_data.get('C3b6', [])[0][1]) and C3b7 >= int(config_data.get('C3b7', [])[0][1]) and C3b8 >= int(config_data.get('C3b8', [])[0][1]) and C3b9 >= int(config_data.get('C3b9', [])[0][1]):
                                                compared_rows = compared_rows._append({
                                                    'Q_ID': key1,
                                                    'T_ID': key2,
                                                    Q_col['FULL_NAME']: C1,
                                                    Q_col['FULL_ADDRESS']: C2,
                                                    Q_col['AGE']: C3b4,
                                                    Q_col['GENDER']: C3b5,
                                                    Q_col['PINCODE']: C3b7,
                                                    Q_col['ICD']: C1b,
                                                    Q_col['RELATIVE']: C3b1,
                                                    'TOTAL': Total_S,
                                                    'PROB': Norm_score
                                                }, ignore_index=True)                            
                                    else:
                                        if C4a >= int(config_data.get('C4a', [])[0][1]) and C4b >= int(config_data.get('C4b', [])[0][1]) and C4c>= int(config_data.get('C4c', [])[0][1]) and C4d >= int(config_data.get('C4d', [])[0][1]) and C4e >= int(config_data.get('C4e', [])[0][1]):
                                            if C4e1 >= int(config_data.get('C4e1', [])[0][1]):
                                                compared_rows = compared_rows._append({
                                                    'Q_ID': key1,
                                                    'T_ID': key2,
                                                    Q_col['FULL_NAME']: C1,
                                                    Q_col['FULL_ADDRESS']: C2,
                                                    Q_col['AGE']: C3b4,
                                                    Q_col['GENDER']: C3b5,
                                                    Q_col['PINCODE']: C3b7,
                                                    Q_col['ICD']: C1b,
                                                    Q_col['RELATIVE']: C3b1,
                                                    'TOTAL': Total_S,
                                                    'PROB': Norm_score
                                                }, ignore_index=True)         
                                            elif C4e2 >= int(config_data.get('C4e2', [])[0][1]):
                                                compared_rows = compared_rows._append({
                                                    'Q_ID': key1,
                                                    'T_ID': key2,
                                                    Q_col['FULL_NAME']: C1,
                                                    Q_col['FULL_ADDRESS']: C2,
                                                    Q_col['AGE']: C3b4,
                                                    Q_col['GENDER']: C3b5,
                                                    Q_col['PINCODE']: C3b7,
                                                    Q_col['ICD']: C1b,
                                                    Q_col['RELATIVE']: C3b1,
                                                    'TOTAL': Total_S,
                                                    'PROB': Norm_score
                                                }, ignore_index=True)         
                                else:
                                    if C5 >= int(config_data.get('C5', [])[0][1]):
                                        if C5a >= int(config_data.get('C5a', [])[0][1]):   
                                            if C5b >= int(config_data.get('C5b', [])[0][1]):
                                                if C5c >= int(config_data.get('C5c', [])[0][1]):
                                                    if C5d >= int(config_data.get('C5d', [])[0][1]):
                                                        if C5e >= int(config_data.get('C5e', [])[0][1]):
                                                            compared_rows = compared_rows._append({
                                                                'Q_ID': key1,
                                                                'T_ID': key2,
                                                                Q_col['FULL_NAME']: C1,
                                                                Q_col['FULL_ADDRESS']: C2,
                                                                Q_col['AGE']: C3b4,
                                                                Q_col['GENDER']: C3b5,
                                                                Q_col['PINCODE']: C3b7,
                                                                Q_col['ICD']: C1b,
                                                                Q_col['RELATIVE']: C3b1,
                                                                'TOTAL': Total_S,
                                                                'PROB': Norm_score
                                                            }, ignore_index=True)                                         
                    pbar.update(1)

    
    Res = "score.xlsx"
    save_dataframe_to_excel(compared_rows, Res)
    score_merger(Res,Res,"Q_ID","T_ID")
    score_remover(Res,Res,"merged")
    sort_excel_descending(Res,Res,"TOTAL")
    extract_matching_records(Res,"QC.xlsx","results.xlsx")
    F_list = ["score.xlsx","QC.xlsx","TC.xlsx","results.xlsx"]
    move_files_to_folder(args.j,F_list)

   
    
if __name__ == "__main__":
    main()