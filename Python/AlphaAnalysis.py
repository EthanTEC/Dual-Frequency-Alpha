import AlphaPlotting as ap
import AlphaAnalysis as aa
import FindFlats as ff
import FindFlats2 as ff2
import DataCollection as dc
import numpy as np

def main():
    try:
        # Load and prepare the data
        tableData, pressureColumns, metaInfo = dc.dataCollect()
    except Exception as e:
        print(f"Error during data collection: {e}")
        return
    
    try:
        # Convert time and date to datetime format
        properTimeDate, _ = ap.plotFromTableData(tableData, pressureColumns, metaInfo, skipPlotting=False)
    except Exception as e:
        print(f"Error during plotting: {e}")
        return
    
    from AlphaGUI import main as gui_main
    col, guiZones = gui_main(properTimeDate, 'Elapsed Time [s]', pressureColumns)
    print(f"User chose column: {col}")
    print(f"User chose zones: {guiZones}")
    # Example usage in your driver:
# --------------------------------
# assume `data` is your DataFrame, with columns:
#    "Elapsed Time [s]"  and  one pressure column, e.g. series = data['well1_press1[psi]']

    time = properTimeDate["Elapsed Time [s]"]
    series = properTimeDate['well1_press1[psi]']

    # Let user pick or compute defaults
    diff_window    = int(input("Derivative smoothing window (pts) [5]: ") or 5)
    diff_threshold = float(input("Transition threshold (abs Δ) [auto]: ") or np.nan)
    if np.isnan(diff_threshold):
        diff_threshold = None

    min_size       = int(input("Minimum zone length (pts) [100]: ") or 100)

    # 1) detect edges
    transitions = ff2.detect_transitions(series, diff_window, diff_threshold)
    print("Detected transitions at indices:", transitions)

    # 2) build zones between them
    zones = ff2.make_step_zones(series, transitions, min_size)
    print("Candidate zones:", zones)

    # 3) plot
    ff2.plot_zones(time, series, zones, title=f"{series.name} (step segmentation)")

    # 4) ask user which to keep
    labels = [str(i) for i in range(1, len(zones)+1)]
    choice = input(f"Select zones to keep {labels}: ")
    chosen = [zones[int(s)-1] for s in choice.split(',') if s.strip() in labels]
    print("You picked:", chosen)


if __name__ == "__main__":
    main()