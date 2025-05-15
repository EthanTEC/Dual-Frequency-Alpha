#!/usr/bin/env python3

"""
Author: [Ethan Predinchuk]
Date: [2025-05-12]
Version: 1.0

This program is designed to analyze pressure data from an Excel file.
It allows the user to select an Excel file, specify the header row, and choose which pressure columns to analyze.
The date must be labelled in the following formate in order to be read correctly:
    Time Column: "Time"
    Pressure Columns: "well1_press1[psi]", "well1_press2[psi]", "well1_press3[psi]", "well1_press4[psi]"
The program will then plot the pressure data over time of the selected columns.

The program has a default cap on the number of pressure columns to be selected, which is set to 4 by default.
To change this, pass a different value to the maxValue parameter in the get_pressure_columns function found in main().
The program also includes error handling for invalid inputs and missing columns.

To change the programs default header row, change the value of the headerRow variable in the main() function.

Future improvements could include:
- Adding more robust error handling for file loading and data processing.
- Implementing a GUI for user input instead of command line prompts.
- Allowing the user to specify the time and presssure column names.
"""

# Import necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
import datetime as dt
import DataCollection as dc

def convertTimeAndDateToDatetime(data, date):
    """
    Convert the time column to datetime format.
    """

    if not pd.api.types.is_datetime64_any_dtype(data["Time"]):
        print("Converting 'Time' column to datetime format.")

        # Convert the date string to a datetime object
        parsedDate = dt.datetime.strptime(date, "%b-%d-%Y").date()
        dateStr = parsedDate.strftime("%Y-%m-%d")
        print(f"Parsed date: {dateStr}")

        # Combine the date with the time column
        data["Time"] = dateStr + " " + data["Time"].astype(str)
        data["Time"] = pd.to_datetime(data["Time"], errors='coerce', format="%Y-%m-%d %H:%M:%S.%f")
    else:
        print("The 'Time' column is already in datetime format.")
    
    return data

def plot_pressure_over_time(data, pressureIndices, testPart):
    """
    Plot pressure over time for each pressure column.
    """


    plt.figure(figsize=(10, 6))
    for column in pressureIndices:
        plt.plot(data["Elapsed Time [s]"], data[column], label=column)
    plt.xlabel('Elapsed Time [s]')
    plt.ylabel('Pressure [psi]')
    plt.title(f'Pressure Over Time of Part: {testPart}')
    plt.legend()
    plt.grid()
    plt.show(block=False)
    input("Press Enter to continue...")

def plotFromTableData(tableData, pressureColumns, metaInfo, skipPlotting=False):
    """
    Plot pressure over time from the table data.
    """
    properTimeDate = convertTimeAndDateToDatetime(tableData, metaInfo.iloc[5, 0].split(": ")[1])
    if properTimeDate is None:
        raise Exception("Failed to convert time and date to datetime format. Check the data format, should be in the form MMM-DD-YYYY.")

    customer = metaInfo.iloc[0, 0].split(": ")[1]
    # Plot pressure over time
    
    elapsedTime = (properTimeDate["Time"] - properTimeDate["Time"].iloc[0]).dt.total_seconds()
    properTimeDate["Elapsed Time [s]"] = elapsedTime
    if not skipPlotting:
        plot_pressure_over_time(properTimeDate,  pressureColumns, customer)
    return properTimeDate, customer

def main():
    while True:
        tableData, pressureColumns, metaInfo = dc.dataCollect()
        if tableData is None or pressureColumns is None:
            print("Failed to load and prepare data.")
            return
        
        plotFromTableData(tableData, pressureColumns, metaInfo)

        continue_choice = input("Do you want to analyze another file? (y/n): ").strip().lower()
        if continue_choice != 'y':
            break
    print("Exiting the program.")

if __name__ == "__main__":
    main()