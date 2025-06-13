import os # operating system interaction
import glob # find pathnames matching specified pattern
import pandas as pd
import numpy as np
import logging
from datetime import datetime
from openpyxl import Workbook, load_workbook # enable opening of and interaction with existing Macro enabled Excel files 
# from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import PatternFill, Alignment, NamedStyle, Font # enable alterations to opened Excel file

def list_excel_files(folder_path, file_pattern, prefix_suffix_separator, date_format):
    '''
    Function to list all files in the specified folder matching the given file pattern.
    
    Parameters:
    folder_path: The path to the folder containing the Excel files.
    file_pattern: The pattern to match the Excel files (e.g., 'IPFFT-*.xlsx').
    prefix_suffix_separator: character used to split prefix and suffix (e.g. '-' or'_')
    date_format: The format the date being searched for is expected to be in (e.g. '%b%y') 
    
    Returns:
    - list of matching file names sorted by date in descending order.
    '''
    # Combine folder path and pattern to form the full search pattern
    search_pattern = os.path.join(folder_path, file_pattern)
    
    # Get list of all matching files (glob used to find pathnames that match specified patterns)
    # https://docs.python.org/3/library/glob.html
    files = glob.glob(search_pattern)
    
    # Raise ValueError if no files matching pattern exists
    if not files:
        raise ValueError("No matching Excel files found in the specified folder.")
    
    # Sort files by extracted date suffix in descending order 
    sorted_files = sorted(files, key=lambda x: datetime.strptime(os.path.basename(x).split(prefix_suffix_separator)
                                                                 [1].split('.')[0], date_format), reverse=True)
    
    return sorted_files

def load_excel_sheet(file_path, sheet_name):
    '''
    Function to load the specified sheet from the given Excel file into a DataFrame.
    
    Parameters:
    file_path: The path to the Excel file.
    sheet_name: The name of the sheet to load.
    
    Returns:
    - df: containing content of the specified sheet.
    '''
    # Read the specified sheet from the Excel file into a DataFrame
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    # Raise an error if the sheet name does not exist
    except ValueError as e:
        raise ValueError(f"Sheet '{sheet_name}' not found in the file.") from e
    
    return df

def validate_column_length(df, columns, lengths):
    '''
    Fucntion to validate that the length of each value in the specified column(s) 
    matches one or more expected lengths.

    Parameters:
    df: DataFrame containing columns to check
    column: Name of the column to check
    lengths: int or list of ints, allowed lengths for the column values
    
    Returns:
    - None: df is simply validated for incorrect data raising an error if incorrect
    '''
    # Convert columns to list if a single string provided
    if isinstance(columns, str):
        columns = [columns]

    # Convert lengths to list if single integer provided
    if isinstance(lengths, int):
        lengths = [lengths]

    # Iterate over each specified column
    for column in columns:
        if column not in df.columns:
            raise KeyError(f"Column '{column}' not found in DataFrame")

        # Check if every value in the column has permitted length
        for idx, value in df[column].items():
            if len(str(value)) not in lengths:
                raise ValueError(f"Row {idx} in column '{column}' contains a value with invalid length.")
    return df

def validate_numeric_columns(df, columns, expected_type):
    '''
    Fucntion to validate that all values in the specified column(s) are either integers or floats,
    depending on the expected type.

    Parameters:
    df: DataFrame containing data to validate
    columns: Name of the column(s) to check
    expected_type: Either 'int' or 'float', as the expected numeric type
    
    Returns:
    - None: df is simply validated for incorrect data raising an error if incorrect
    '''

    # Ensure columns is a list
    if isinstance(columns, str):
        columns = [columns]

    # Check that the expected_type is either 'int' or 'float'
    if expected_type not in ['int', 'float']:
        raise TypeError(f"Invalid expected_type '{expected_type}'. Must be 'int' or 'float'.")

    # Iterate over all specified columns
    for column in columns:
        # Ensure the column exists in the DataFrame
        if column not in df.columns:
            raise KeyError(f"Column '{column}' not found in DataFrame.")

        # Check each value in the column
        for idx, value in df[column].items():
            if expected_type == 'int' and not isinstance(value, int):
                raise TypeError(f"Row {idx} in column '{column}' contains a non-integer value.")
            elif expected_type == 'float' and not isinstance(value, (float, int)): # Allow int for float check
                raise TypeError(f"Row {idx} in column '{column}' contains a non-float value.")
    
    return df
    
def get_cell_content_as_string(source_dataframe, source_row, source_col):
    '''
    Function to combine a specified string with the contents of a dataframe cell to enable dynamic update.
    
    Parameters:
    source_dataframe: The name of the DataFrame that contains the source text.
    source_row: Row index of the DataFrame. 
    source_col: Column index of the DataFrame. 
   
    Returns:
    - cell_content: the value within the cell specified cast as a string. 
    '''
    # Raise KeyError if column name does not exist in the DataFrame.
    if source_col not in source_dataframe.columns:
        raise KeyError(f"Column '{source_col}' not found in the DataFrame")

    # Raise IndexError if the row index is outside the DataFrame range.
    if source_row >= len(source_dataframe):
        raise IndexError(f"Row index {source_row} is out of range")
    
    # Extract cell content from the DataFrame and cast as a string
    cell_content = str(source_dataframe.at[source_row, source_col])
    
    return cell_content

def map_fft_period(periodname, yearnumber):
    """
    Fucntion to generate the FFT_Period based on mapping 'Periodname' to the correct year based on 'Yearnumber'.

    Parameters:
    periodname: The period name (e.g., 'JANUARY', 'FEBRUARY').
    yearnumber: The year range (e.g., '2024-25').

    Returns:
    - fft_period: The corresponding FFT_Period (e.g., 'Jan-25').
    """
    # Define a dictionary for month abbreviations
    month_abbrev = {
        'JANUARY': 'Jan', 'FEBRUARY': 'Feb', 'MARCH': 'Mar',
        'APRIL': 'Apr', 'MAY': 'May', 'JUNE': 'Jun',
        'JULY': 'Jul', 'AUGUST': 'Aug', 'SEPTEMBER': 'Sep',
        'OCTOBER': 'Oct', 'NOVEMBER': 'Nov', 'DECEMBER': 'Dec'
    }

    # Raise ValueError if periodname is invalid
    if periodname not in month_abbrev:
        raise ValueError(f"Invalid period name '{periodname}'.")
    # Raise ValueError if yearnumber is invalid
    if not (len(yearnumber) == 7 and yearnumber[4] == '-'):
        raise ValueError("Yearnumber mot int the correct format.")

    # Extract the start and end years from the 'Yearnumber' (e.g., from '2024-25')
    start_year = yearnumber[:4] # First 4 characters represent the start year
    end_year = yearnumber[5:] # Characters after the hyphen represent the end year

    # Determine which part of the year to use based on the period name
    if periodname in ['JANUARY', 'FEBRUARY', 'MARCH']:
        fft_period = f"{month_abbrev[periodname]}-{end_year}" # Use the end year
    else:
        fft_period = f"{month_abbrev[periodname]}-{start_year[2:]}" # Use the start year

    return fft_period

def remove_columns(df, columns_to_remove):
    '''
    Function to remove columns not required from the DataFrame including helper columns added during suppression processing.

    Parameters:
    df: DataFrame form which to remove columns.
    columns_to_remove: single or list of columns to be removed (e.g. ['Period', 'Response Rate'] from Ward Level submissions).
    
    Returns:
    - df: with specified field(s) removed.
    '''
    # Raise a KeyError if column(s) for removal not in DataFrame
    missing_columns = [col for col in columns_to_remove if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns are not in the DataFrame: {missing_columns}")

    # drop defined columns
    return df.drop(columns=columns_to_remove)

def rename_columns(df, new_column_names):
    '''
    Function to rename DataFrame columns to align with final product as specificed by stakeholders.
    
    Parameters:
    df: DataFrame requiring column name changes.
    new_column_names: dictionary of columns to be renamed (e.g. {'STP Code': 'ICB_Code', 'STP Name': 'ICB_Name'} from Ward Level submissions).
    
    Returns:
    - df: with specified fields renamed.
    '''
    # Check and store any columns for renaming not present in the DataFrame and raise KeyError
    missing_columns = [col for col in new_column_names.keys() if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns to be renamed do not exist in the DataFrame: {missing_columns}")

    # rename columns
    return df.rename(columns=new_column_names)

def replace_non_matching_values(df, column_name, target_value, replacement_value):
    '''
    Function to replace values in a specified column of a DataFrame that do not match a target value with a replacement value.
    
    Parameters:
    df: The DataFrame being worked on.
    column_name: The name of the column to check and replace values.
    target_value: The value that should be retained in the column.
    replacement_value: The value to replace non-matching values with.
    
    Returns:
    - df: with non-matching values replaced in the specified column
    '''
    # Raise a KeyError if the specified column name is not in the DataFrame
    if column_name not in df.columns:
        raise KeyError(f"Column '{column_name}' not found in the DataFrame")

    # Replace any value in the specified column that does not match the target_value (e.g. if not IS1 then NHS)
    df[column_name] = df[column_name].apply(lambda x: x if x == target_value else replacement_value)
    
    return df

def sum_grouped_response_fields(df, columns_to_group_by):
    '''
    Function to aggregate data submitted at one level to data at a higher level e.g. aggregating Trust level data to ICB level data.
    This allows Total Response, Total Eligible and Breakdown of Repsonses fields to be aggregated to the necessary level before new
    calculations are carried out to generate Percentage Positive and Percentage Negative fields at the new level.
    
    Parameters:
    df: DataFrame requiring grouping and summing.
        
    Returns:
    - df: with selected content aggregated by specified fields and all numerical fields summed.
    '''    
    # Check and store any columns specified not present in the DataFrame and raise KeyError
    missing_columns = [col for col in columns_to_group_by if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns are missing in the DataFrame: {missing_columns}")
    
    return df.groupby(columns_to_group_by).sum().reset_index()

def create_data_totals(df, current_fft_period, total_column_name, columns_to_sum):
    '''
    Function to create monthly data totals row from existing DataFrame. Totals will have a specified column set to 'Total', 
    'Period' will be the same as the current FFT Period, and all other specified columns will be the sum of the 
    corresponding columns.

    Parameters:
    df: The DataFrame containing the original data.
    current_fft_period: period defined relating to the current source input file
    total_column_name: The name of the column that will have 'Total' as the value in the new row.
    columns_to_sum: A list of column names that should be summed.

    Returns:
    - df: containing single row of totals for specified columns.
    '''
    # Check and store any columns specified not present in the DataFrame and raise KeyError
    missing_columns = [col for col in columns_to_sum if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns are missing in the DataFrame: {missing_columns}")
    
    # Create a dictionary to hold the new total row data and set the specified column to 'Total'
    total_data = {total_column_name: 'Total'}
    
    # Assume 'Period' is the same for all rows and take the value from the first row
    total_data['Period'] = current_fft_period
    
    # Iterate over the list of columns to sum and calculate the sum for each
    for col in columns_to_sum:
        total_data[col] = df[col].sum()
    
    # Create a DataFrame for the total row using the total_data dictionary
    total_df = pd.DataFrame([total_data])
  
    # Return the modified DataFrame with the total row added
    return total_df

def append_dataframes(df1, df2):
    '''
    Function to append one DataFrame to another.

    Parameters:
    df1: DataFrame containing one set of data.
    df2: DataFrame containing another set of data with aligned columns.

    Returns:
    - df: with content of both source DataFrames appended into one DataFrame.
    '''
    # Append second DataFrame to the first DataFrame
    return pd.concat([df1, df2], ignore_index=True)

def create_percentage_field(df, percentage_column, sum_column_one, sum_column_two, total_columm):
    '''
    Function to generate percentage positive and negative columns for survey results using percentage calculation
    sum_columns / total * 100   
    
    Parameters:
    df: DataFrame with fields to calculate for addition of percentage columns.
    percentage_column: new column added containing results of percentage calculation.
    sum_column_one: first of two columns summed to esablish percentage of total.
    sum_column_two: second of two columns summed to esablish percentage of total.
    total_column: totaa response field used to establish the percentage the two other columns represent.
            
    Returns:
    - df: with new percentage column added.
    '''        
    # Check and store any columns specified not present in the DataFrame and raise KeyError 
    missing_columns = [col for col in [sum_column_one, sum_column_two, total_columm] if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns are missing in the DataFrame: {missing_columns}")

    df[percentage_column] = round(((df[sum_column_one] + df[sum_column_two]) / df[total_columm]) * 1, 2)
    
    return df

def remove_rows_by_cell_content(df, column_name, cell_value):
    '''
    Function to remove rows from DataFrame based on specified cell content in a given column.
    
    Parameters:
    df: DataFrame from which rows will be removed.
    column_name: The name of the column where the specific value is located.
    cell_value: The specific value that will determine row removal.
    
    Returns:
    - df: with specified rows removed.
    '''
    # Raise KeyError if column_name for searching values not in DataFrame
    if column_name not in df.columns:
        raise KeyError(f"Column '{column_name}' not found in the DataFrame")
    
    # Display the initial number of rows in the DataFrame before removal
    logging.info(f"Initial number of rows: {len(df)}")
    
    # Filter DataFrame keeping rows where specified column content does NOT match specified cell value.
    df = df[df[column_name] != cell_value]
    
    # Display number of rows after filtering helping identify number removed
    logging.info(f"Number of rows after removal: {len(df)}")
    
    return df

def reorder_columns(df, column_order):
    '''
    Function to reorder the DataFrame columns based on a specified order.
    
    Parameters:
    df: The DataFrame to reorder.
    column_order: List of desired column order.
    
    Returns:
    - DataFrame with columns reordered according to provided list.
    '''
    
    # Ensure the specified order includes all columns from the DataFrame
    if set(column_order) != set(df.columns):
        raise ValueError("The columns list must contain exactly the same columns as the DataFrame.")
    
    # Reorder the columns of the DataFrame
    df = df[column_order]
    
    return df

def convert_fields_to_object_type(df, fields_to_convert):
    '''
    Function converting specified columns to object type (e.g. like those undergoing suppression process). When saved into Excel this ensures number values 
    remain as numbers, while avoiding raising python errors when suppressing values with non-numeric values. Object data type stores any 
    data type. - https://pandas.pydata.org/pandas-docs/stable/user_guide/basics.html#basics-dtypes

    Parameters:
    df: The input DataFrame.
    fields_to_convert: List of fields that require converting e.g. for number (int) fields where suppression applied to string (str) 

    
    Returns:
    - df: specified columns converted to object type.
    '''

    # Convert specified fields to avoid data type incompatibility error 
    df[fields_to_convert] = df[fields_to_convert].astype(object)
    return df

def replace_missing_values(df, replacement_value):
    '''
    Function to replace null (NaN) values in a DataFrame with specified value.

    Parameters:
    df: DataFrame to be cleaned.

    Returns:
    - df: Cleaned DataFrame with NaN values replaced.
    '''
    df_cleaned = df.copy()
    # Replace nan with specified value
    with pd.option_context('future.no_silent_downcasting', True):
        df_cleaned.fillna(replacement_value, inplace=True)
    
    return df_cleaned

def count_nhs_is1_totals(df, code_column_name, is1_count_var_name, nhs_count_var_name):
    '''
    Function counting number of rows with 'IS1' and 'NHS' in a specified column, storing counts in variables with 
    specified names for use as summary totals.

    Parameters:
    df: containing the data.
    code_column_name: The name of the column to check for 'IS1' and 'NHS' values.
    is1_count_var_name: The name of the variable to store the count of 'IS1'.
    nhs_count_var_name: The name of the variable to store the count of 'NHS'.

    Returns:
    - Dictionary with counts stored against specified variable names.
    '''
    # Raise KeyError if the code_column_name does not exist in the DataFrame
    if code_column_name not in df.columns:
        raise KeyError(f"Column '{code_column_name}' not found in the DataFrame.")
    
    # Count number of rows containing 'IS1'
    count_of_IS1 = df[df[code_column_name] == 'IS1'].shape[0]

    # Count number of rows containing 'NHS'
    count_of_NHS = df[df[code_column_name] == 'NHS'].shape[0]

    # Return counts in a dictionary with specified variable names
    return {is1_count_var_name: count_of_IS1, nhs_count_var_name: count_of_NHS}

def add_dataframe_column(df, column_name, column_value):
    '''
    Function to create a new DataFrame column, specifying the DataFrame, column name and value it has.

    Parameters:
    df: DataFrame to which the new column will be added.
    column_name: The name of the new column.
    column_value: The value with which to populate or instantiate the new column.

    Returns:
    - df: with the new populated column added.
    '''
    # Rasie TypeError if new column name not specified as a string or list
    if isinstance(column_name, str):
        column_name = [column_name]    
    
    if not isinstance(column_name, list):
        raise TypeError("New column name should be a list or a string.")

    # Add the new column to the DataFrame with the specified initial value
    df[column_name] = column_value
    
    return df

def add_submission_counts_to_df(df, code_column_name, is1_count, nhs_count, target_column):
    '''
    Function to add submission counts as specified column to the specified DataFrame.

    Parameters:
    df: The DataFrame to which the counts should be added.
    code_column_name: The name of the column containing 'IS1' and 'NHS'.
    is1_count: The count of rows with 'IS1' taken from a previous aggregation.
    nhs_count: The count of rows with 'NHS' taken from a previous aggregation.
    target_column: The name of the column to add counts to.

    Returns:
    - df: updated with tatget column populated with counts.
    '''
    # Raise KeyError if `code_column_name` or `target_column` does not exist in the DataFrame.
    if code_column_name not in df.columns or target_column not in df.columns:
        raise KeyError(f"One or both of columns '{code_column_name}' or '{target_column}' are missing from the DataFrame.")

    # Add the count to the row where ICB Code is 'IS1'
    df.loc[df[code_column_name] == 'IS1', target_column] = is1_count

    # Add the count to the row where ICB Code is 'NHS'
    df.loc[df[code_column_name] == 'NHS', target_column] = nhs_count

    return df

def update_monthly_rolling_totals(df1, df2, current_fft_period):
    '''
    Function to transfer data from df1 to df2 in the next blank row or overwrite an existing row if the period exists.
    
    Parameters:
    df1: Source DataFrame containing data to be transferred (must have specific columns and format).
    df2: Destination DataFrame where the data will be transferred (must have specific columns and format).
    current_fft_period: value generated for the current fft period.
    
    Returns:
    - df2: updated with data from df1 transferred to the next available row if the period doesn't already exist, 
    or overwritten if it does.
    '''
    # Raise KeyError if necessary columns are missing from df1.
    required_columns_df1 = ['Submitter Type', 'Number of organisations submitting', 'Total Responses', 'Percentage Positive', 'Percentage Negative']
    missing_columns_df1 = [col for col in required_columns_df1 if col not in df1.columns]
    if missing_columns_df1:
        raise KeyError(f"Missing columns in df1: {missing_columns_df1}")
    
    # Raise KeyError if FFT Period column is missing from df2.
    if 'FFT Period' not in df2.columns:
        raise KeyError("'FFT Period' column is missing in df2.")

    # Create a new row for df2 based on df1 data
    new_row = {
        'FFT Period': current_fft_period,
        'Total submitters': df1.loc[df1['Submitter Type'] == 'Total', 'Number of organisations submitting'].values[0],
        'Number of NHS submitters': df1.loc[df1['Submitter Type'] == 'NHS', 'Number of organisations submitting'].values[0],
        'Number of Independent submitters': df1.loc[df1['Submitter Type'] == 'IS1', 'Number of organisations submitting'].values[0],
        'Total responses to date': 0, # Initiating with 0
        'Total NHS responses to date': 0, # Initiating with 0
        'Total independent responses to date': 0, # Initiating with 0
        'Monthly total responses': df1.loc[df1['Submitter Type'] == 'Total', 'Total Responses'].values[0],
        'Monthly NHS responses': df1.loc[df1['Submitter Type'] == 'NHS', 'Total Responses'].values[0],
        'Monthly independent responses': df1.loc[df1['Submitter Type'] == 'IS1', 'Total Responses'].values[0],
        'Monthly total percentage positive': df1.loc[df1['Submitter Type'] == 'Total', 'Percentage Positive'].values[0],
        'Monthly NHS percentage positive': df1.loc[df1['Submitter Type'] == 'NHS', 'Percentage Positive'].values[0],
        'Monthly independent percentage positive': df1.loc[df1['Submitter Type'] == 'IS1', 'Percentage Positive'].values[0],
        'Monthly total percentage negative': df1.loc[df1['Submitter Type'] == 'Total', 'Percentage Negative'].values[0],
        'Monthly NHS percentage negative': df1.loc[df1['Submitter Type'] == 'NHS', 'Percentage Negative'].values[0],
        'Monthly independent percentage negative': df1.loc[df1['Submitter Type'] == 'IS1', 'Percentage Negative'].values[0]
    }
    
    # Check if 'current_fft_period' already exists in 'FFT Period' column of df2
    if current_fft_period in df2['FFT Period'].values:
        # Find the index of the existing row
        row_index = df2[df2['FFT Period'] == current_fft_period].index[0]
        
        # Overwrite the existing row with the new values
        for key, value in new_row.items():
            df2.at[row_index, key] = value
        
        logging.info(f"Figures for the current month ({current_fft_period}) have been overwritten.")
    else:
        # Convert the dictionary to a DataFrame and add it to bottom of df2
        new_row_df = pd.DataFrame([new_row])
        df2 = pd.concat([df2, new_row_df], ignore_index=True)
        
        logging.info(f"New figures for the current month ({current_fft_period}) have been added to the DataFrame.")
    
    return df2

def update_cumulative_value(df, first_column, second_column):
    '''
    Function takes the value from last row of a specified source column, adds it to the value from the second-to-last row of 
    another specified source column, and updates the target column with the result in the last row.

    Parameters:
    df: The DataFrame containing the data.
    source_column: The column from which to take the value from the bottom row.
    target_column: The column where the result should be added to and updated in the second-to-bottom row.
    
    Returns:
    - df: updated with the cumulative value in the last row.
    '''
    # Raise KeyError if either column does not exist in the DataFrame
    if first_column not in df.columns or second_column not in df.columns:
        raise KeyError(f"One or both columns '{first_column}' and '{second_column}' do not exist in the DataFrame.")

    # Raise ValueError if there are at less than two rows
    if len(df) < 2:
        raise ValueError("The DataFrame must have at least two rows to update cumulative values.")
    
    # Identify the index of the last and second-to-last rows
    last_idx = df.index[-1] # Index of the last row
    second_last_idx = df.index[-2] # Index of the second-to-last row
    
    # Extract the values from the specified columns
    first_value = df.at[last_idx, first_column]
    second_value = df.at[second_last_idx, second_column]
    
    # Addition values together
    cumulative_value = first_value + second_value
    
    # Update the target column in the second-to-bottom row with the new value
    df.at[last_idx, second_column] = cumulative_value
    
    # Return the updated DataFrame
    return df

def update_existing_excel_sheet(file_path, sheet_name, updated_df):
    '''
    Function to update a specific sheet in an existing Excel file with new data from a DataFrame.
    
    Parameters:
    file_path: Path to the existing Excel file.
    sheet_name: Name of the sheet to update.
    updated_df: The DataFrame containing the updated data to be written back to the sheet.
    
    Returns:
    None - Specified Excel sheet updated.
    '''
    # Raise FileNotFoundError if the file does not exist
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file '{file_path}' does not exist.")

    # Load the existing Excel file using openpyxl engine
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Write the updated DataFrame to the specified sheet
        updated_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    logging.info(f"Sheet '{sheet_name}' in '{file_path}' has been updated successfully.")

def copy_value_between_dataframes(df_source, df_target, source_column, source_row, target_column, target_row):
    '''
    Function to copy a value from a specific row and column in one DataFrame to a specific row and column in another DataFrame.

    Parameters:
    df_source: The source DataFrame from which the value will be copied.
    df_target: The target DataFrame where the value will be pasted.
    source_column: The column in the source DataFrame from which to copy the value.
    source_row: The row index in the source DataFrame from which to copy the value.
    target_column: The column in the target DataFrame where the value will be pasted.
    target_row: The row index in the target DataFrame where the value will be pasted.

    Returns:
    - df: with the value pasted into the specified row and column.
    '''
    # Raise KeyError if either source or target columns do not exist in respective DataFrames
    if source_column not in df_source.columns:
        raise KeyError(f"Source column '{source_column}' does not exist in df_source.")
    if target_column not in df_target.columns:
        raise KeyError(f"Target column '{target_column}' does not exist in df_target.")

    # Raise IndexError if specified rows in the source or target DataFrames do not exist
    if source_row not in df_source.index:
        raise IndexError(f"Source row '{source_row}' does not exist in df_source.")
    if target_row not in df_target.index:
        raise IndexError(f"Target row '{target_row}' does not exist in df_target.")
    
    # Copy the value from the source DataFrame
    value_to_copy = df_source.at[source_row, source_column]
    
    # Paste the value into the target DataFrame
    df_target.at[target_row, target_column] = value_to_copy
    
    return df_target

def new_column_name_with_period_prefix(period_prefix, new_column_suffix):
    '''
    Function to create a new column name for use in creating a new column, with the prefix of a specified data period.
    
    Parameters:
    period_prefix: month-year variable e.g., current_fft_month, previous_fft_month required for new column name.
    new_column_suffix: suffix required of the new column name.
    
    Returns:
    - new column name
    '''
    #Raise TypeError if either period_prefix or new_column_suffix is not a string.
    if not isinstance(period_prefix, str):
        raise TypeError("period_prefix must be a string")
    
    if not isinstance(new_column_suffix, str):
        raise TypeError("new_column_suffix must be a string")
    
    # Join new_column_suffix wih period_prefix to form full column name
    new_column_name = period_prefix + '_' + new_column_suffix
    
    return new_column_name

def sort_dataframe(df, df_fields, directions):
    '''
    Function to sort the selected DataFrame by specified field(s).
    
    Parameters:
    df: DataFrame to sort by.
    df_fields: field(s) to sort DataFrame by. Can be single field (string) or list of fields.
    directions: sort direction(s). (True for A to Z, False for Z to A). Can be single boolean or a list for each field.
    
    Returns:
    - df: with content ordered by the specified field(s).
    (ignore_index (True) rests the index to 0-based sequence rather than retaining the excisting indexing)
    '''
    # Ensure df_fields exist in the DataFrame
    if isinstance(df_fields, str):
        df_fields = [df_fields]
    if not all(field in df.columns for field in df_fields):
        # Raise KeyError if any listed sort fields are not in the DataFrame
        raise KeyError("One or more fields to sort by do not exist in the DataFrame.")

    # Sort the DataFrame based on the provided fields and directions
    df = df.sort_values(by=df_fields, ascending=directions, ignore_index=True)
    return df

def create_first_level_suppression(df, first_level_suppression, responses_field):
    '''
    Function to add first level suppression field to ensure any row of data reporting less than 5 total responses is marked for suppression.
    This also applies to first level suppression of next level aggregation e.g. at Trust level this applies to ICB, 
    at Site Level it applies to Trust, and at Ward Level it applies to Site.
    
    Parameters: 
    df: The input DataFrame.
    responses_field: the field containing the survey response totals e.g. Responses.
    first_level_suppression: the name of the new field to add containing first level suppression .   
    
    Returns: 
    - df: with additional column distinguishing columns requiring direct (first level) suppression (1) versus not (0).
    '''
    # Raise KeyError if responses_field does not exist
    if responses_field not in df.columns:
        raise KeyError(f"'{responses_field}' does not exist in the DataFrame.")
    
    # Create the direct suppression column based on condition applied to the responses column (value is 1 if less than 5, otherwise value is 0)
    df[first_level_suppression] = df[responses_field].apply(lambda x: 1 if 0 < x <5 else 0) 
    return df
    
def create_icb_second_level_suppression(df, first_level_suppression, second_level_suppression):
    '''
    Function to add second level suppression column for icb level only to ensure for any icb with first level suppression, an additoinal icb is suppressed
    maximising security against submissions being patient identifiable.
    
    Parameters: 
    df: The input DataFrame.
    first_level_suppression: the field showing rows where first level suppression has been added (as 1).
    second_level_suppression: the new field to be added to highlight rows where second level suppression is required (1) or not (0).
    
    Returns: 
    - df: with additional column distinguishing columns requiring second level suppression (1) versus not (0).
    '''
    # Create the second level suppression field with default value of 0
    df[second_level_suppression] = 0
    
    # Iterate through the DataFrame starting from the 2nd row (index 1) as the first row can't have a preceeding row being suppressed
    for i in range(1, len(df)):
        # Check if the previous rows 'first_level_suppression' is 1
        if df.at[i-1, first_level_suppression] == 1:
            # Set second_level_suppression to 1 for the current row
            df.at[i, second_level_suppression] = 1
    return df    

def confirm_row_level_suppression(df, suppression_field, *suppression_columns):
    '''
    Function to add a new suppression field which shows 1 if the row needs suppressing based on need for suppression in any of the other fields checking 
    need for suppression, and 0 if it doesn't need suppressing. 

    Parameters:
    df: The input DataFrame.
    suppression_field: new field added to show where rows will need suppressing.
    *suppression_columns: fields to check for suppression.

    Returns:
    - df: with a additional field confirming which rows need suppressing based on all other suppression fields.
    '''
    # Check and store any columns specified not present in the DataFrame and raise KeyError 
    missing_columns = [col for col in suppression_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following suppression columns are missing from the DataFrame: {', '.join(missing_columns)}")
    
    # Initialize the overall suppression_field column with 0s
    df[suppression_field] = 0
 
    # Iterate over all specified suppression columns
    for column in suppression_columns:
        # Update the 'suppresison_field' column if any specified suppression column has a 1
        df[suppression_field] = df[suppression_field] | (df[column] == 1)
    
    # Convert 'suppress_field' to integer (whole number) type 
    df[suppression_field] = df[suppression_field].astype(int)
    
    return df

def suppress_data(df, overall_suppression_field, first_level_suppression_field):
    '''
    Function replaces specified column values with '*' based on conditions of 'first_level_suppression' and 'overall_suppression_field'. 

    Parameters:
    df: The input DataFrame
    overall_suppression_field: field highlighting rows in the DataFrame that require some level of suppression.
    first_level_suppression_field: field highlighting rows in the DataFrame that require percentage fields suppressing as well.

    Returns:
    - df: modified with suppressed values.
    '''
    # List of columns to replace with '*' if 'overall suppression required
    response_column_suppression = ['Very Good', 'Good', 'Neither Good nor Poor', 'Poor', 'Very Poor', 'Dont Know']

    # List of additional columns to replace with '*' if 'first_level_suppression' is also 1
    percentage_column_suppression = ['Percentage Positive', 'Percentage Negative']

    # Check and store any specified columns not present in the DataFrame and raise KeyError
    all_suppression_columns = response_column_suppression + percentage_column_suppression
    missing_columns = [col for col in all_suppression_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(f"The following columns are missing in the DataFrame: {', '.join(missing_columns)}")

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        if row[overall_suppression_field] == 1:
            # Replace specified columns with '*' for 'overall level suppression columns'
            df.loc[index, response_column_suppression] = '*'
            if row[first_level_suppression_field] == 1:
                # Replace percentage columns with '*' where 'first_level_suppression' required
                df.loc[index, percentage_column_suppression] = '*'
    return df

def move_independent_provider_rows_to_bottom(df):
    '''
    Function to move all indpendent provider (IS1) rows to the bottom of the dataframe.

    Parameters:
    df: The input DataFrame.

    Returns:
    - df: modified with suppressed values.
    '''
    # Create a boolean mask identifying rows where 'ICB Code' contains independent provider marker 'IS1'
    is_ip = df['ICB Code'].str.contains('IS1')

    # Separate the DataFrame into two parts: rows with 'IS1' and rows without 'IS1'
    ip_rows = df[is_ip] # Rows where 'ICB Code' contains 'IS1'
    other_rows = df[~is_ip] # Rows where 'ICB Code' does not contain 'IS1'

    # Concatenate the DataFrame with 'other_rows' first and 'ip_rows' second
    result_df = pd.concat([other_rows, ip_rows], ignore_index=True)

    return result_df

def adjust_percentage_field(df, percentage_column):
    '''
    Function to convert existing percentage positive and negative columns to ensure they present correctly in final outputs
    (percentage_column / 100)
    
    Parameters:
    df: DataFrame with fields to calculate for adjusting of percentage columns.
    percentage_column: existing column containing percentage in incorrect format.

    Returns:
    - df: with percentage column adjusted.
    '''        
    # Raise KeyError if the specified percentage column does not exist in the DataFrame.
    if percentage_column not in df.columns:
        raise KeyError(f"Column '{percentage_column}' not found in the DataFrame.")

    df[percentage_column] = round(df[percentage_column] / 100, 2)
    
    return df

def rank_organisation_results(df, org_field, responses_field, rank_field):
    '''
    Function to add Rank column to the DataFrame, ranking Responses within each e.g. Site Code group.

    Parameters:
    df: The input DataFrame
    org_field: the field containing the organisation level code to sort by e.g. Site_Code.
    responses_field: the field containing the survey response totals e.g. Responses.
    rank_field: the name of the new field to add containing ranking e.g. Rank.

    Returns:
    - df: with an additional column containing rankings by organisation code. Any Responses value of 0 will be ranked 0.
    All remaining responses by organisation code will be ranked from 1 for lowest non 0 response level, upwards.    
    '''
    # Raise KeyError if either org_field or responses_field do not exist
    if org_field not in df.columns or responses_field not in df.columns:
        raise KeyError(f"'{org_field}' or '{responses_field}' does not exist in the DataFrame.")
    
    # Create a mask for non-zero responses
    non_zero_mask = df[responses_field] != 0

    # Create the ranking column with default of 0 for all values
    df[rank_field] = 0

    # Apply ranking to non-zero responses within each org_field grouping excluding rows where repsonses are 0
    df.loc[non_zero_mask, rank_field] = df[non_zero_mask].groupby(org_field)[responses_field] \
                                                  .rank(method='dense') \
                                                  .astype(int)
    return df

def create_second_level_suppression(df, first_level_suppression, rank_field, second_level_suppression):
    '''
    Function to add second level suppression column to ensure where an organisation has more than one submission row, 
    where the first submission is greater than 0 but less than 5 and requires suppression, 
    the second submission is also suppressed maximising security against submissions being patient identifiable.
    
    Parameters: 
    df: The input DataFrame.
    rank_field: the field showing ranked order for an organisations submissions.
    first_level_suppression: the field showing rows where first level suppression has been added (as 1).
    second_level_suppression: the new field to be added to highlight rows where second level suppression is required (1) or not (0).
    
    Returns: 
    - df: with additional column distinguishing columns requiring second level suppression (1) versus not (0).
    '''
    # Raise KeyError if either first_level_suppression or rank_field do not exist
    if first_level_suppression not in df.columns or rank_field not in df.columns:
        raise KeyError(f"'{first_level_suppression}' or '{rank_field}' does not exist in the DataFrame.")

    # Create the second level suppression field with default value of 0
    df[second_level_suppression] = 0
    
    # Iterate through the DataFrame starting from the 2nd row (index 1) as first row can't contain a 2 in rank_field
    for i in range(1, len(df)):
        # Check if 'Rank' is 2 and the previous rows 'first_level_suppression' is 1
        if df.at[i, rank_field] == 2 and df.at[i-1, first_level_suppression] == 1:
            # Set second_level_suppression to 1 for the current row
            df.at[i, second_level_suppression] = 1
    return df

def add_suppression_required_from_upper_level_column(upper_level_df, lower_level_df, upper_level_suppression_column, code_lookup_field, suppression_lookup_field):
    '''
    Function adding a new column to DataFrame with value of 1 for rows where upper level output shows suppression is required against 
    lower level orgnaisations/sites/wards, resulting in suppression where 'Rank' column in the lower level DataFrame is 1 
    (if it features for the Business Level Code) or 1 and 2 (if they both feature for the Business Level Code).
    Example - if an ICB Code is flagged as requiring suppression in the ICB level output, in the Trust level output, with all Trusts sorted in Rank order 
    by ICB, the first Trust reporting more than 0 responses (the Rank 1 Trust for the ICB), and the Rank 2 Trust (if there is one) will be marked for suppresison 
    at Trust level. 

    Parameters:
    upper_level_df: DataFrame containing data aggregated to upper level where identifier codes are unique.
    lower_level_df: DataFrame containing data aggregated to lower level where identifier codes present in 
                    upper level are not unique, where the new column will be added for suppression.
    upper_level_suppression_column: new field added to lower_level_df showing need for suppression.
    code_lookup_field: field from upper_level_df containing the organisation site/ward codes for dictionary key.
    suppression_lookup_field: field from upper_level_df containing suppression status for dictionary value.

    Returns:
     - lower_level_df: with the new suppression column added aligned with need to suppress from upper level DataFrame.
    '''
    # Raise KeyError if either code_lookup_field or suppression_lookup_field do not exist in the upper_level_df
    if code_lookup_field not in upper_level_df.columns or suppression_lookup_field not in upper_level_df.columns:
        raise KeyError(f"Missing required columns '{code_lookup_field}' or '{suppression_lookup_field}' in upper_level_df.")

    # Raise KeyError if either code_lookup_field or Rank do not exist in the upper_level_df
    if code_lookup_field not in lower_level_df.columns or 'Rank' not in lower_level_df.columns:
        raise KeyError(f"Missing required columns '{code_lookup_field}' or 'Rank' in lower_level_df.")

    # Create a dictionary from 'upper_level_df' with code_lookup_field as the key and suppression_lookup_field as the value
    suppression_dict = upper_level_df.set_index(code_lookup_field)[suppression_lookup_field].to_dict()
    
    # Initialize the new column in 'lower_level_df' with 0s
    lower_level_df[upper_level_suppression_column] = 0

    # Iterate over each Business Level Code in lower_level_df
    for business_level_code in lower_level_df[code_lookup_field].unique():
        # Check if suppression is required for this Business Level Code based on upper_level_df
        if suppression_dict.get(business_level_code, 0) == 1:
            # Filter the rows with the current Business Level Code
            icb_rows = lower_level_df[lower_level_df[code_lookup_field] == business_level_code]

            # Assign suppression for Rank 1 if it exists for the business level code
            if any(icb_rows['Rank'] == 1):
                lower_level_df.loc[(lower_level_df[code_lookup_field] == business_level_code) & (lower_level_df['Rank'] == 1), upper_level_suppression_column] = 1

            # Assign suppression for Rank 2 if it exists for the business level code
            if any(icb_rows['Rank'] == 2):
                lower_level_df.loc[(lower_level_df[code_lookup_field] == business_level_code) & (lower_level_df['Rank'] == 2), upper_level_suppression_column] = 1

    return lower_level_df

def join_dataframes(df1, df2, on='column_to_join_on', how='left', validate='one_to_one'):
    '''
    Function to join one DataFrame with another (side by side). 
    
    Parameters:
    df1: DataFrame containing one set of data.
    df2: DataFrame containing another set of data with aligned columns.    
    
    Return:
    - df: with content of both source DataFrames joined into one DataFrame 
    '''
    # Raise KeyError if the join `on` column is missing from either DataFrame.
    if on not in df1.columns or on not in df2.columns:
        raise KeyError(f"Join column '{on}' not found in one of the DataFrames.")
    
    # Raise ValueError if an invalid join 'how' type is specified    
    valid_how = ['left', 'right', 'inner', 'outer']
    if how not in valid_how:
        raise ValueError(f"Invalid join type '{how}'.")

    df = df1.join(df2.set_index(on), on, how, validate)
    
    return df

def replace_character_in_columns(df, columns, target_chars, replacement_char):
    """
    Function to replace specific character(s) with another character in the specified DataFrame column(s).
    
    Parameters:
    df: DataFrame where replacement is required.
    columns: Single column or list of columns where replacemented might be needed.
    target_chars: The character(s) to search for in the specified column(s).
    replacement_char: The character to replace the target character.
    
    Returns:
    - df: modified with the specified character(s) replaced.
    """
    
    # If 'columns' is a single string (i.e., one column), convert it to a list for uniform handling
    if isinstance(columns, str):
        columns = [columns]

    # Loop through each column
    for column in columns:
        # Ensure the column exists
        if column in df.columns:
            # Loop through each target character in the target_chars list
            for char in target_chars:
                # Replace the target character with the replacement character in the specified column
                df[column] = df[column].astype(str).str.replace(char, replacement_char, regex=False)
        else:
            # Raise a KeyError if the specified column name is not in the DataFrame
            raise KeyError(f"Column '{column}' not in DataFrame")
    
    return df

def remove_duplicate_rows(df):
    '''
    Function remove duplicate rows in a DataFrame to retain only rows of unique values. 

    Parameters:
    df: DataFrame where duplicate rows require removing.
    
    Returns:
    - df: with any duplicate rows removed.
    '''    
    
    return df.drop_duplicates()

def limit_retained_columns (existing_df, columns):
    '''
    Function to retain column/columns of an existing dataframe. 
    Useful where the columns to retain are fewer than the columns to remove. 

    Parameters:
    existing_df: DataFrame form which to filter columns to retain.
    columns: single or list of columns to be retained (e.g. ['ICB Code', 'ICB Name'] from Ward Level submissions).
    
    Returns:
    - df: with just the specified field(s) retained.
    '''    
    # If 'columns' is a single string (i.e., one column), convert it to a list for uniform handling
    if isinstance(columns, str):
        columns = [columns]

    # Raise TypeError if columns is not a string or list
    if not isinstance(columns, list):
        raise TypeError("columns should be a string or list of strings")

    # Raise KeyError if column(s) for retaining not in DataFrame
    missing_columns = [col for col in columns if col not in existing_df.columns]
    if missing_columns:
        raise KeyError(f"Required columns are missing form the DataFrame.")
       
    return existing_df.filter(columns, axis=1)

def open_macro_excel_file(source_file_path):
    '''
    Function to open existing Macro-Enabled Excel file where one exists to be used as a template to generate Output file from.
    
    Parameters:
    source_file_path: Path to the existing macro-enabled Excel file.
    
    Returns:
    - workbook: Loaded as object.    
    '''
    # Load the existing macro-enabled workbook retaining vba
    workbook = load_workbook(source_file_path, keep_vba=True)    

    return workbook

def write_dataframes_to_sheets(workbook, dfs_info):
    '''
    Function to write DataFrames to specified sheets and cells in a workbook already opened as an object.
    
    Parameters:
    workbook: The workbook object where DataFrames will be written.
    dfs_info: Dataframes information written in tuples. For each sheet the tuple will contain 
    (DataFrame, Excel sheet name, start row, start column).
    
    Returns:
    - None: Workbook object is updated with content from Dataframes based on Tuples.
    '''
    # Iterate over the DataFrames and corresponding sheet and cell positions
    for df, sheet_name, start_row, start_col in dfs_info:
        # Select the sheet
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
        else:
            raise ValueError(f"Sheet {sheet_name} does not exist in the workbook.")
        # Iterate over DataFrame rows (i) and columns (j) to write data to Excel without headers
        for i, row in enumerate(df.itertuples(index=False, name=None), start=start_row):
            for j, value in enumerate(row, start=start_col):
                sheet.cell(row=i, column=j).value = value

def update_cell_with_formatting(workbook, sheet_name, start_row, start_col, data, font_size=10, bg_color="FFFFFF", bold=True, font_name="Verdana", 
                                align_horizontal="center", align_vertical="center"):
    '''
    Function to update a specific cell or range of cells with data and apply formatting.
    
    Parameters:
    workbook: The workbook object where the sheet resides.
    sheet_name: Name of the sheet to update.
    start_row: The starting row index (number/integer) for updating.
    start_col: The starting column index (number/integer) for the updating.
    data: Data to paste into the cell or range of cells.
    font_size: Font size (number/integer) for the cell(s).
    bg_color: Background color for the cell(s) in hex format (e.g., 'FFFFFF' for white; 'BFBFBF' for grey).
    bold: 
    font_name: Font name for the cell(s) (e.g., "Calibri").
    align_horizontal: Horizontal alignment (e.g., "center", "left", "right").
    align_vertical: Vertical alignment (e.g., "center", "top", "bottom").
    
    Returns:
    - None: Workbook object is updated with specified content.
    '''
    # Select the sheet
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        raise ValueError(f"Sheet {sheet_name} does not exist in the workbook.")
    
    # Check if the data is a list of lists (indicating multiple rows/columns)
    if isinstance(data, list):
        for i, row_data in enumerate(data):
            for j, value in enumerate(row_data):
                cell = sheet.cell(row=start_row + i, column=start_col + j)
                cell.value = value
                # Apply formatting
                cell.font = Font(size=font_size, bold=bold, name=font_name)
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                cell.alignment = Alignment(horizontal=align_horizontal, vertical=align_vertical)
    else:
        # Single cell update
        cell = sheet.cell(row=start_row, column=start_col)
        cell.value = data
        # Apply formatting
        cell.font = Font(size=font_size, bold=bold, name=font_name)
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell.alignment = Alignment(horizontal=align_horizontal, vertical=align_vertical)

def create_percentage_style(workbook):
    '''
    Function to create a percentage style with 0 decimal places and add it to the workbook. NameStyles can 
    only be defined and registered to a workbook once, so defininition needs to be destrinct from application. 

    Parameters:
    workbook: The openpyxl workbook object.

    Returns:
    - percentage_style: The created NamedStyle.
    '''
    # Raise TypeError if the the workbook is invalid object
    if not isinstance(workbook, Workbook):
        raise TypeError("The 'workbook' must be an openpyxl Workbook object.")

    # Check if the percentage style already exists to avoid re-creating it
    for style in workbook.named_styles:
        # Ensure we are dealing with NamedStyle objects and not strings
        if isinstance(style, NamedStyle) and style.name == "percentage_style":
            return style # If style exists, return it

    # Create the percentage style if it does not exist
    percentage_style = NamedStyle(name="percentage_style")
    percentage_style.number_format = '0%' # Set number format for percentage with 0 decimal places

    # Register the new style to the workbook
    workbook.add_named_style(percentage_style)

    return percentage_style

def format_column_as_percentage(workbook, sheet_name, start_row, start_cols, percentage_style):
    '''
    Function to format specified Excel sheet column(s) as percentages with 0 decimal places.

    Parameters:
    workbook: The workbook object where the sheet resides.
    sheet_name: The name of the sheet to update.
    start_row: The starting row index (number/integer) for foramtting.
    start_col: List of starting column indecies (number/integer) for formatting.
    
    Returns:
    - None: Workbook object is updated with specified formatting.
    '''
    
    # Select the sheet
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        raise ValueError(f"Sheet {sheet_name} does not exist in this workbook.")
   
    # Apply formatting (percentage format with 0 decimal places) to specified columns starting from the specified row
    for start_col in start_cols:
        # Loop from the starting row to the last row
        for row in range(start_row, sheet.max_row + 1): 
            cell = sheet.cell(row=row, column=start_col)
            # Apply the percentage format with 0 decimal places
            cell.style = percentage_style 

def combine_text_and_dataframe_cells(input_text_or_cell_1, input_text_or_cell_2):
    '''
    Function to combine a specified string with the contents of a dataframe cell to enable dynamic update.
    
    Parameters:
    input_text_or_cell_1: First text string or variable with cell content.
    input_text_or_cell_2: Second text string or variable with cell content.
   
    Returns:
    - combined_string: the text string and cell content combined into a string for further use. 
    '''
        
    # Combine the input text or cell content
    combined_string = f"{input_text_or_cell_1} {input_text_or_cell_2}"
    
    return combined_string

def save_macro_excel_file(workbook, source_file_path, new_folder_path, new=False, prefix=None, fft_period_suffix=None):
    '''
    Save the workbook, either replacing the existing file or saving it as a new file with a specified name.
    
    Parameters:
    workbook: The workbook object to be saved.
    source_file_path: The path of the existing macro-enabled Excel file. Retain name used for loading in file.
    new_folder_path: The folder path assigned for saving a new macro-enabled Excel file.
    new: Boolean - If True, saves the file as a new file with a name generated by prefix and suffix. 
         If False, it saves to the original source file path.
    prefix: Prefix for the name of the Excel file.
    fft_period_suffix = the fft period for which the output is being saved e.g., current_fft_period
    
    Returns:
    - None: Excel file saved to specified file path. 
    '''
    if new:
        # Construct the new file name
        new_file_name = f"{prefix}-{fft_period_suffix}.xlsm"
        # Asign new file location
        new_file_path = os.path.join(new_folder_path, new_file_name)        
        # Save as a new file
        workbook.save(new_file_path)
        logging.info(f"Workbook saved as {new_file_path}")
    else:
        # Replace the existing file
        workbook.save(source_file_path)
        logging.info(f"Workbook saved as {source_file_path}")






# def check_current_month_in_rolling_totals(current_fft_period, df2):
    '''
    Function to check if the current period already exists in df2.
    
    Parameters:
    current_fft_period: value generated for the current fft period.
    df2: Destination DataFrame where the data would be checked against (must have specific columns and format).
    
    Returns:
    - message stating if the period does not exist in df2.
    '''
    
    # Check if the period from df1 already exists in df2
#    period = current_fft_period
#    if period in df2['FFT Period'].values:
#        raise ValueError(f"The period '{period}' already exists in df2.")
#    else:
#        return 'Current FFT Period not yet added to Monthly Rolling Totals.'
