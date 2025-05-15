#!/usr/bin/env python3

import pandas as pd
import tkinter as tk
from tkinter import filedialog


def select_excel_file():
    """
    Open a file dialog to select an Excel file.
    Returns the selected file path.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path

def load_excel_data(file_path, headerRow = 7):
    """
    Load the Excel file into a pandas DataFrame.
    The header row is specified by the user.
    Returns the DataFrame.
    """
    try:
        data = pd.read_excel(file_path, header=int(headerRow)-1)
        return data
    except Exception as e:
        print(f"Error loading file: {e}")
        return None

def get_non_negative_integers(
    prompt="Enter a non-negative integer or press Enter to skip: ",
    allowBlank=True,
    multiple=False,
    maximum=None):
    """
    Get non-negative integers from user input.
    This function can handle both single and multiple inputs, and takes input in the form of a comma-separated string.
    Continues to prompt the user until valid input is received.
    Returns a list of integers or a single integer, or an empty string if allowBlank is True.
    """
    while True:
        if not multiple:
            user_input = input(prompt).strip()

            if user_input == "":
                if allowBlank:
                    return ""
                else:
                    raise Exception("Blank input is not allowed. Please enter a non-negative integer.")

            if user_input.isdigit():
                value = int(user_input)
                if maximum is not None and value > maximum:
                    raise Exception(f"Value {value} exceeds the maximum allowed value of {maximum}.")
                return value
            else:
                raise Exception("Invalid input. Please enter a non-negative integer" +
                      (" or leave blank." if allowBlank else "."))
        
        else:
            user_input = input(prompt).strip()

            if user_input == "":
                if allowBlank:
                    return []
                else:
                    raise ValueError("Blank input is not allowed. Please enter non-negative integers.")

            parts = user_input.split(',')
            selected_ints = []
            

            for val in parts:
                val = val.strip()
                if not val.isdigit():
                    raise ValueError("Invalid input. Please enter comma-separated non-negative integers.")
                    
                num = int(val)
                if maximum is not None and num > maximum:
                    raise ValueError(f"Value {num} exceeds the maximum allowed value of {maximum}.")

                selected_ints.append(num)

            print(f"Selected input: {selected_ints}")
            return selected_ints


def get_pressure_columns(maxValue = 4, looping= True):
    """
    Get the pressure columns from user input.
    The user can specify multiple columns separated by commas.
    The maximum number of columns is specified by maxValue.
    Returns a list of integers representing the selected columns.
    """
    
    while True:
        try:
            userInput = get_non_negative_integers(
                prompt=f"Enter the pressure column numbers separated by commas to a maximum of {maxValue} (e.g., 1,2,3). Ensure columns have the title well1_press#[psi]: ",
                allowBlank=False,
                multiple=True,
                maximum=maxValue)
            
            # Parse and clean input
            selectedColumns = [int(x) for x in userInput]
            
            # Remove duplicates and return
            unique_columns = sorted(set(selectedColumns))
            return unique_columns

        except ValueError as ve:
            print(f"Input error: {ve}")
            if not looping:
                print("Exiting due to input error.")
                break
        except Exception as e:
            print(f"Unexpected error: expected one value, exitting: {e}")
            break

def get_header_row(headerRow = 7, looping= True):
    """
    Get the header row from user input.
    Default header row is 7.
    Returns the header row number.
    """
    while True:
        try:
            headerRow = get_non_negative_integers(
                prompt="Enter the header row number (default is 7): ")
            if headerRow == "":
                return 7
            else:
                return int(headerRow)
        except Exception:
            print("Invalid input. Please enter a non-negative integer.")
            if not looping:
                print("Exiting due to input error.")
                break
        except ValueError as e:
            print(f"Unexpected error, expected multiple values, exitting: {e}")
            break

def dataCollect():
    # File path to the Excel file
    file_path = select_excel_file()
    if not file_path:
        raise Exception("No file selected.")
    
    # Load the data
    headerRow = get_header_row()
    
    data = load_excel_data(file_path, headerRow)
    if data is None:
        raise Exception("Failed to load data from the Excel file.")

    # Take users desired pressure columns for analysis
    pressureColumns = get_pressure_columns() # To change the default number of columns to be selected, add/change the maxValue parameter in this function

    # Collect the important data from the table in the excel file
    tableData = data.iloc[6:, :].copy()
    expectedPressureColumns = [f"well1_press{i}[psi]" for i in pressureColumns]
    
    missingColumns = [col for col in expectedPressureColumns if col not in tableData.columns]
    if missingColumns:
        print(f"Warning: The following columns were not found in the data: {missingColumns}")
        proceed = input("Do you want to continue with the available columns? (y/n): ").strip().lower()
        if proceed != 'y':
            print("Operation cancelled by user due to missing columns.")
            exit()
    
    # Check if the time column is in the correct format
    if headerRow > 0:
        metaInfo = pd.read_excel(file_path, nrows=headerRow-1, header=None)
        if metaInfo is None:
            raise Exception("Failed to load metadata from the Excel file.")
    else:
        print(" No metadata found in the file.")

    return tableData, expectedPressureColumns, metaInfo