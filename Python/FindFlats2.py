import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

def detect_transitions(series, diff_window=5, diff_threshold=None):
    """
    series: pd.Series of pressure values
    diff_window: compute rolling-mean on abs(diff) over this many points
    diff_threshold: if None, defaults to 3× median of that rolling-diff
    
    Returns a sorted list of integer indices where a “step” occurs.
    """
    # 1) raw abs derivative
    deriv = series.diff().abs()
    # 2) smooth it
    smooth = deriv.rolling(window=diff_window, min_periods=1).mean()
    # 3) choose threshold
    if diff_threshold is None:
        diff_threshold = smooth.median() * 3
    # 4) find where it spikes
    edges = smooth[smooth > diff_threshold].index.to_numpy()
    # drop duplicates and very-close points
    cuts = []
    last = -np.inf
    for idx in edges:
        if idx - last > diff_window:
            cuts.append(idx)
            last = idx
    return cuts

def make_step_zones(series, transitions, min_size=20):
    """
    Given your ordered transition indices [i1,i2,i3...], build
    zones = [(0, i1), (i1, i2), (i2, i3), ..., (last, len)] and
    drop any shorter than min_size.
    """
    zones = []
    all_pts = [0] + list(transitions) + [len(series)]
    for a, b in zip(all_pts, all_pts[1:]):
        if (b - a) >= min_size:
            zones.append((a, b))
    return zones

def plot_zones(times, series, zones, title=None):
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