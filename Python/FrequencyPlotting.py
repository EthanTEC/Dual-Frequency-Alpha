#!/usr/bin/env python3

def perform_frequency_analysis(data, time_column, pressure_columns):
    """
    Perform frequency analysis on the pressure data.
    """
    time = data[time_column]
    dt = (time.iloc[1] - time.iloc[0]).total_seconds()  # Assuming time is in datetime format
    sampling_rate = 1 / dt

    for column in pressure_columns:
        pressure = data[column]
        n = len(pressure)
        freq = np.fft.fftfreq(n, d=dt)
        fft_values = fft(pressure)
        
        plt.figure(figsize=(10, 6))
        plt.plot(freq[:n // 2], np.abs(fft_values[:n // 2]))
        plt.xlabel('Frequency (Hz)')
        plt.ylabel('Amplitude')
        plt.title(f'Frequency Analysis - {column}')
        plt.grid()
        plt.show()