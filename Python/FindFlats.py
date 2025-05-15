#!/usr/bin/env python3

import matplotlib.pyplot as plt
import pandas as pd

def get_user_thresholds(default_window=25, default_tol=25.0, default_min_size=500):
    """
    Prompt the user for:
      • window size for rolling-std
      • std-dev threshold
      • minimum flat-zone length (in points)
    Returns (window_size:int, tol:float, min_size:int).
    """
    while True:
        try:
            w = input(f"Enter window size for rolling-std (default {default_window}): ").strip()
            window_size = int(w) if w else default_window
            if window_size < 1:
                raise ValueError("Window size must be ≥ 1.")

            t = input(f"Enter max std-dev threshold (default {default_tol}): ").strip()
            tol = float(t) if t else default_tol
            if tol < 0:
                raise ValueError("Threshold must be ≥ 0.")

            m = input(f"Enter minimum flat-zone length in points (default {default_min_size}): ").strip()
            min_size = int(m) if m else default_min_size
            if min_size < 1:
                raise ValueError("Minimum size must be ≥ 1.")

            return window_size, tol, min_size
        except ValueError as e:
            print(f"Invalid input: {e}")


def detect_flat_mask(series, window_size, tol, min_periods=1):
    """
    Compute rolling std-dev with min_periods=1 so the very first points
    can contribute, then threshold.
    """
    rolling_std = series.rolling(window=window_size, min_periods=min_periods).std()
    return rolling_std.le(tol).fillna(False)

def clean_mask(mask, max_hole=5, max_island=5):
    """
    - fill any False runs of length <= max_hole  (closing small gaps)
    - drop any True runs of length <= max_island (remove tiny islands)
    Returns a new pandas Series of bools.
    """
    arr = mask.to_numpy().astype(int)
    N = len(arr)

    # 1) fill small holes
    i = 0
    while i < N:
        if arr[i] == 1:
            j = i
            while j < N and arr[j] == 1:
                j += 1
            # now [i, j) is a True block
            # look at the next False run
            k = j
            while k < N and arr[k] == 0:
                k += 1
            hole_length = k - j
            if hole_length <= max_hole:
                arr[j:k] = 1
            i = k
        else:
            i += 1

    # 2) remove small islands
    i = 0
    while i < N:
        if arr[i] == 0:
            j = i
            while j < N and arr[j] == 0:
                j += 1
            # now [i, j) is a False block; skip
            i = j
        else:
            # a True block
            j = i
            while j < N and arr[j] == 1:
                j += 1
            island_length = j - i
            if island_length <= max_island:
                arr[i:j] = 0
            i = j

    return pd.Series(arr.astype(bool), index=mask.index)

def extract_zones(mask):
    """
    From a boolean mask, extract list of (start_idx, end_idx)
    pairs for each contiguous True region.
    """
    zones = []
    in_zone = False
    for i, ok in enumerate(mask):
        if ok and not in_zone:
            start = i
            in_zone = True
        elif not ok and in_zone:
            zones.append((start, i))
            in_zone = False
    if in_zone:
        zones.append((start, len(mask)))
    return zones


def filter_zones_by_size(zones, min_size):
    """
    Remove any zone whose length (end - start) < min_size.
    """
    return [(s, e) for (s, e) in zones if (e - s) >= min_size]


def plot_flat_zones(times, series, zones, title=None):
    """
    Plot the series, highlight each zone in red, and label 1..N.
    """
    plt.figure(figsize=(12, 6))
    plt.plot(times, series, label=series.name)

    for idx, (start, end) in enumerate(zones, 1):
        xs = times.iloc[start:end]
        ys = series.iloc[start:end]
        plt.plot(xs, ys, color='red', linewidth=2)
        mid = start + (end - start) // 2
        plt.text(times.iloc[mid], series.iloc[mid], str(idx),
                 ha='center', va='top',
                 bbox=dict(boxstyle="round,pad=0.2", fc="yellow"))

    if title:
        plt.title(title)
    plt.xlabel("Elapsed Time [s]")
    plt.ylabel(series.name)
    plt.grid(True)
    plt.legend()
    plt.show(block=False)
    plt.pause(0.001)


def select_zones(zones):
    """
    Ask user to pick which labeled zones (1..N) to keep.
    Returns list of zone tuples.
    """
    if not zones:
        print("No flat zones detected.")
        return []
    labels = [str(i) for i in range(1, len(zones) + 1)]
    prompt = f"Select zones to continue with ({','.join(labels)}), e.g. 1,3: "
    while True:
        sel = input(prompt).split(',')
        chosen = []
        for s in sel:
            s = s.strip()
            if s in labels:
                chosen.append(zones[int(s) - 1])
        if chosen:
            return chosen
        print("Invalid selection; try again.")


def find_flats(data, pressure_columns):
    """
    Master function.  
    1) Gets window_size, tol, min_size from user.  
    2) Detects & extracts zones for each pressure column.  
    3) Filters out zones shorter than min_size.  
    4) Plots each; then asks “Happy with these? (y/n)”.  
       If no, repeats step 1.  
    5) Once accepted, prompts user to pick which zones to keep.  
    Returns dict: { column_name: [ (start,end), … ] }
    """
    while True:
        window_size, tol, min_size = get_user_thresholds()
        all_zones = {}

        # detect → extract → filter
        for col in pressure_columns:
            ser = data[col].astype(float)
            raw_mask = detect_flat_mask(ser, window_size, tol, min_periods=1)
            cleaned = clean_mask(raw_mask, max_hole=window_size//2, max_island=window_size//10)
            zones = extract_zones(cleaned)
            zones = filter_zones_by_size(zones, min_size)
            all_zones[col] = zones

        # plot & review
        for col, zones in all_zones.items():
            plot_flat_zones(
                data["Elapsed Time [s]"],
                data[col],
                zones,
                title=f"{col} (win={window_size}, tol={tol}, minpts={min_size})"
            )

        yn = input("Are you happy with these flat zones? (y/n): ").strip().lower()
        if yn == 'y':
            break
        else:
            plt.close('all')

    # final selection
    selected = {}
    for col, zones in all_zones.items():
        print(f"\nColumn: {col}")
        chosen = select_zones(zones)
        selected[col] = chosen

    return selected