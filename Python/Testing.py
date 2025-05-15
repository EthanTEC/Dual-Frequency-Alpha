
import DataCollection as dc
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import datetime as dt

def convertTimeAndDateToDatetime(data, date):
    """
    Convert the time column to datetime format.
    """
    print(f"Data before conversion: {data.head()}")

    if not pd.api.types.is_datetime64_any_dtype(data["Time"]):
        print("Converting 'Time' column to datetime format.")

        # Convert the date string to a datetime object
        parsedDate = dt.datetime.strptime(date, "%b-%d-%Y").date()
        dateStr = parsedDate.strftime("%Y-%m-%d")

        if dateStr is None:
            # Use default date if parsing fails
            dateStr = "2025-01-01"

        # Combine the date with the time column
        data["Time"] = dateStr + " " + data["Time"].astype(str)
        data["Time"] = pd.to_datetime(data["Time"], errors='coerce', format="%Y-%m-%d %H:%M:%S")
        # data["Time"] = pd.to_datetime(data["Time"], errors='coerce', format="%Y-%m-%d %H:%M:%S.%f")
    else:
        print("The 'Time' column is already in datetime format.")
    
    return data

def plot_pressure_over_time(data, pressureIndices, testPart, lower, upper):
    """
    Plot pressure over time for each pressure column.
    """
    # print(f"length of data: {len(data)}")
    # plt.figure(figsize=(10, 6))
    # for column in pressureIndices:
    #     plt.plot(data["Elapsed Time [s]"], data[column], label=column)
    # plt.xlabel('Elapsed Time [s]')
    # plt.ylabel('Pressure [psi]')
    # plt.title(f'Pressure Over Time of Part: {testPart}')
    # plt.legend()
    # plt.grid()
    # plt.show(block=False)

    plt.figure(figsize=(10, 6))
    for column in pressureIndices:
        plt.plot(data["Elapsed Time [s]"][lower:upper], data[column][lower:upper], label=column)
    plt.xlabel('Elapsed Time [s]')
    plt.ylabel('Pressure [psi]')
    plt.title(f'Pressure Over Time of Part: {testPart}')
    plt.legend()
    plt.grid()
    plt.show(block=False)
    input("Press Enter to continue...")

def plotFromTableData(tableData, pressureColumns, metaInfo, lower, upper, skipPlotting=False):
    """
    Plot pressure over time from the table data.
    """
    properTimeDate = convertTimeAndDateToDatetime(tableData, metaInfo.iloc[5, 0].split(": ")[1])
    if properTimeDate is None:
        raise Exception("Failed to convert time and date to datetime format. Check the data format, should be in the form MMM-DD-YYYY.")

    customer = metaInfo.iloc[0, 0].split(": ")[1]
    # Plot pressure over time

    print(f"ProperTimeDate: {properTimeDate.head()}")

    elapsedTime = (properTimeDate["Time"] - properTimeDate["Time"].iloc[0]).dt.total_seconds()

    print(f"Elapsed Time: {elapsedTime.head()}")
    properTimeDate["Elapsed Time [s]"] = elapsedTime
    if not skipPlotting:
        plot_pressure_over_time(properTimeDate,  pressureColumns, customer, lower, upper)
    return properTimeDate, customer

def freqAnalysis(data, pressureColumns, metaInfo, lower, upper):
    """
    Perform frequency analysis on the pressure data.
    """

    df_slice = data.iloc[lower:upper]  # note: iloc is end-exclusive, so use 1001 to include row 1000

    # extract time and pressure arrays
    t = df_slice['Elapsed Time [s]'].interpolate(method='linear', limit_direction='both').to_numpy()
    y = df_slice[pressureColumns[0]].interpolate(method='linear', limit_direction='both').to_numpy()


    y_demeaned = y - np.mean(y)  # remove DC offset

    # compute sampling interval (assuming roughly uniform spacing)
    dt = np.mean(np.diff(t))
    fs = 1.0 / dt             # sampling frequency
    N = y_demeaned.size                # number of samples

    # perform FFT
    Y = np.fft.fft(y_demeaned)
    freqs = np.fft.fftfreq(N, d=dt)

    # only keep the positive frequencies
    mask = freqs >= 0
    freqs_pos = freqs[mask]
    amp_spectrum = np.abs(Y[mask]) * 2.0 / N   # scaled amplitude

    # plot
    plt.figure(figsize=(8,4))
    plt.plot(freqs_pos, amp_spectrum)
    plt.xlabel('Frequency [Hz]')
    plt.ylabel('Amplitude')
    plt.title('FFT of wellPressure')
    plt.grid(True)
    plt.show(block=False)

def main():
    
    try:
        # Load and prepare the data
        tableData, pressureColumns, metaInfo = dc.dataCollect()
    except Exception as e:
        print(f"Error during data collection: {e}")
        return
    
    firstRun = True

    while True:

        if firstRun:
            lower = 0
            upper = len(tableData) - 1

        else:
            lower = int(input("Enter the lower bound for plotting: "))
            upper = int(input("Enter the upper bound for plotting: "))

        try:
            # Convert time and date to datetime format
            properTimeDate, _ = plotFromTableData(tableData, pressureColumns, metaInfo, lower, upper, skipPlotting=False)
        except Exception as e:
            print(f"Error during plotting: {e}")
            return
        
        if not firstRun:
            freqAnalysis(properTimeDate, pressureColumns, metaInfo, lower, upper)
        
        firstRun = False

        continue_choice = input("Do you want to analyze another file? (y/n): ").strip().lower()
        if continue_choice != 'y':
            break

if __name__ == "__main__":
    main()