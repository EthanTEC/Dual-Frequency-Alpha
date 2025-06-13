import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox as tkmsg
import customtkinter as ctk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import RectangleSelector
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backends.backend_pdf import PdfPages
from PIL import Image, ImageTk
from datetime import datetime
from json import load

# ──────────────────────────────────────────────────────────────────────────────
# 1) VERSION AND UPDATE_INFO_URL (DEPRECATED)
# ──────────────────────────────────────────────────────────────────────────────
__version__ = "1.4.0"
UPDATE_INFO_URL = "https://raw.githubusercontent.com/EthanTEC/Dual-Frequency-Alpha/main/Python/update_info.json"

def try_delete_old_exe():
    """
    If launched with "--replace-old <old_path>", wait briefly then delete <old_path>.
    This allows the new EXE to overwrite/delete the previous version during updates.
    """
    args = sys.argv
    if "--replace-old" in args:
        idx = args.index("--replace-old")
        if idx + 1 < len(args):
            old_path = args[idx + 1]
            time.sleep(1.0)  # Give OS time to release file lock
            try:
                if os.path.isfile(old_path):
                    os.remove(old_path)
            except Exception:
                pass
            del sys.argv[idx: idx + 2]

# Always attempt to delete old EXE before anything else (COMMENT BACK IN TO REINITIALIZE INSTALLER)
# try_delete_old_exe()

# Determine BASE_PATH for bundled executable (PyInstaller) or dev environment
if getattr(sys, "frozen", False):
    BASE_PATH = sys._MEIPASS
else:
    script_dir = os.path.abspath(os.path.dirname(__file__))
    BASE_PATH = os.path.abspath(os.path.join(script_dir, os.pardir))


# Enable high-DPI awareness on Windows
try:
    if sys.platform.startswith('win'):
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

class AlphaAnalysisApp(ctk.CTk):
    """
    Main application class for Alpha Analysis.
    Provides:
      - Loading from Excel (interactive header selection).
      - Loading from a cached Parquet file (faster than Excel for large datasets).
      - Drawing flat zones interactively.
      - Time and frequency domain plots per zone.
      - Saving analysis/plots to PDF, or saving raw DataFrame to Parquet.
      - Version checking and update mechanism.
    """

    def __init__(self):
        """
        Initialize CTk window, create control panel and plotting area,
        and set up internal state.
        """
        super().__init__()
        self.title("")

        # Base dimensions for font scaling
        self.base_width = 1600
        self.base_height = 900
        self.base_font_size = 12
        self.ui_font = "Segoe UI"
        self._setup_scaling()

        # Configure grid: controls panel (col 0), plot area (col 1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Containers
        self.control_container = ctk.CTkFrame(self)
        self.control_container.grid(row=0, column=0, sticky="ns")
        self.control_container.grid_propagate(False)

        self.plot_container = ctk.CTkFrame(self)
        self.plot_container.grid(row=0, column=1, sticky="nsew")

        # Scrollable controls
        self._setup_control_canvas()

        # Internal state
        self.df = None
        self.zones = []
        self.time_col = None
        self.pressure_cols = []
        self.elapsed_col = None
        self.test_date = None
        self.header_row = None
        self.collected_date_event = threading.Event()
        self.bad_date_event = threading.Event()
        self.elapsed_mode = tk.BooleanVar(value=False)
        self.save_data_mode = tk.BooleanVar(value=False)

        # Build UI
        self._build_controls()
        self._build_plot()

        # Resize debounce
        self._resize_job = None
        self.bind("<Configure>", self._on_configure)
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Comment back in to reinitialize autoupdate check at the start of the program
        # self._check_for_updates(autoUpdating=True)

    def _setup_scaling(self):
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        scale = min(screen_w / self.base_width, screen_h / self.base_height)
        self.tk.call('tk', 'scaling', scale)
        new_size = max(6, min(int(self.base_font_size * scale), 20))
        self.ui_style = (self.ui_font, new_size)

    def _setup_control_canvas(self):
        self.control_canvas = tk.Canvas(self.control_container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.control_container, orient="vertical", command=self.control_canvas.yview)
        self.control_frame = ctk.CTkFrame(self.control_canvas)

        self.control_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.control_canvas.grid(row=0, column=0, sticky="nsew")
        self.control_container.grid_rowconfigure(0, weight=1)
        self.control_container.grid_columnconfigure(0, weight=1)

        self.control_window = self.control_canvas.create_window((0, 0), window=self.control_frame, anchor="nw")
        self.control_frame.bind("<Configure>", 
                                lambda e: self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all")))

    def _build_controls(self):
        """
        Construct all widgets in the control panel:
          1) Buttons to load Excel or load saved Parquet.
          2) Interactive header-row preview (Treeview).
          3) Dropdowns/listboxes to select time and pressure columns.
          4) Controls for minimum zone size, confirm zones, save options, and update check.
        """
        cf = self.control_frame
        cf.grid_columnconfigure(0, weight=1)
        r = 0

        # 1. Load Data
        ctk.CTkLabel(cf, text="1. Load Data", font=self.ui_style).grid(row=r, column=0, sticky="w", pady=(5,2))
        r+=1
        self.browse_btn = ctk.CTkButton(cf, text="Browse Excel/Parquet...", command=self._browse_file, font=self.ui_style)
        self.browse_btn.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1
        self.file_lbl = ctk.CTkLabel(cf, text="No file chosen", wraplength=200, font=self.ui_style)
        self.file_lbl.grid(row=r, column=0, sticky="w", pady=2)
        r+=1

        # 2. Header Selection
        self.hdr_lbl = ctk.CTkLabel(cf, text="Header row: None", font=self.ui_style)
        self.hdr_lbl.grid(row=r, column=0, sticky="w", pady=2)
        self.hdr_lbl.grid_remove()  # hide initially
        r+=1
        self.preview = tk.Frame(cf, height=150)
        self.preview.grid(row=r, column=0, sticky="ew", pady=2)
        self.preview.grid_propagate(False)
        self.preview.grid_remove()  # hide initially
        r+=1
        self.tree = ttk.Treeview(self.preview, show="headings", height=5)
        vs = ttk.Scrollbar(self.preview, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self.preview, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vs.set, xscroll=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.preview.grid_rowconfigure(0, weight=1)
        self.preview.grid_columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_header_select)
        r+=0  # keep r same until switch placed

        # Elapsed switch
        self.elapsed_switch = ctk.CTkSwitch(cf, text="Use Elapsed Only", variable=self.elapsed_mode, font=self.ui_style)
        self.elapsed_switch.grid(row=r, column=0, sticky="w", pady=2)
        r+=1

        # 3. Select Columns
        ctk.CTkLabel(cf, text="3. Select Columns", font=self.ui_style).grid(row=r, column=0, sticky="w")
        r+=1
        ctk.CTkLabel(cf, text="Time Column:", font=self.ui_style).grid(row=r, column=0, sticky="w")
        r+=1
        self.time_cb = ttk.Combobox(cf, state="disabled")
        self.time_cb.grid(row=r, column=0, sticky="ew", pady=2)
        self.time_cb.bind("<<ComboboxSelected>>", lambda e: setattr(self, 'time_col', self.time_cb.get()))
        r+=1
        ctk.CTkLabel(cf, text="Pressure Columns:", font=self.ui_style).grid(row=r, column=0, sticky="w")
        r+=1
        self.p_list = tk.Listbox(cf, selectmode="multiple", height=4, font=self.ui_style)
        self.p_list.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1

        # 4. Load & Plot
        self.load_btn = ctk.CTkButton(cf, text="4. Load & Plot (Excel)", command=self._load_data_thread, font=self.ui_style)
        self.load_btn.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1
        ctk.CTkLabel(cf, text="Min Zone Size (s):", font=self.ui_style).grid(row=r, column=0, sticky="w")
        r+=1
        self.min_var = tk.DoubleVar(value=30.0)
        self.min_entry = ctk.CTkEntry(cf, textvariable=self.min_var, font=self.ui_style)
        self.min_entry.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1

        # 5. Confirm Zones
        self.confirm_btn = ctk.CTkButton(cf, text="5. Confirm Zones", command=self._confirm, font=self.ui_style)
        self.confirm_btn.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1

        # 6. Save Options
        ctk.CTkLabel(cf, text="Save Options", font=self.ui_style).grid(row=r, column=0, sticky="w", pady=(10,2))
        r+=1
        self.save_data_switch = ctk.CTkSwitch(cf, text="Save as data (Parquet)", variable=self.save_data_mode, font=self.ui_style)
        self.save_data_switch.grid(row=r, column=0, sticky="w", pady=2)
        r+=1

        # 7. Save Button
        self.save_btn = ctk.CTkButton(cf, text="6. Save", command=self._save_analysis, font=self.ui_style)
        self.save_btn.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1

        # 8. Export Zones
        self.export_zones_btn = ctk.CTkButton(cf, text="7. Export Zones to Parquet", command=self._export_zones, font=self.ui_style)
        self.export_zones_btn.grid(row=r, column=0, sticky="ew", pady=2)
        r+=1

        # --- Section 8: Check for Updates (Comment back in to readd update button to allow users to check update url for new updates)---
        # self.update_btn = ctk.CTkButton(
        #     self.control, text="8. Check for Updates", command=self._check_for_updates, font=self.ui_style
        # )
        # self.update_btn.pack(fill="x", pady=2)

    def _build_plot(self):
        """
        Set up the matplotlib Figure and Axes inside a CTkFrame, attach the canvas,
        and initialize the RectangleSelector (for drawing zones) plus loading GIF/logo.
        """
        self.fig, self.ax = plt.subplots()
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_container)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.canvas, self.plot_container)
        self.canvas.mpl_connect("button_press_event", self._on_click)

        self.rs = None # RectangleSelector will be instantiated once data is loaded

        # Loading GIF setup
        loading_widget = self.canvas.get_tk_widget()
        canvas_bg = loading_widget.cget("bg")
        self.loading_gif_path = os.path.join(BASE_PATH, "Images", "LoadingGIF.gif")
        self.loading_gif_frames = []
        self.current_frame = 0
        self.loading_label = tk.Label(
            self.plot_container, bd=0, bg=canvas_bg, highlightthickness=0
        )
        self.finished_loading_event = threading.Event()

        # Logo setup (top-right corner)
        logo_widget = self.canvas.get_tk_widget()
        canvas_bg = logo_widget.cget("bg")
        self.logo_path = os.path.join(BASE_PATH, "Images", "TEC.jpg")
        logo = Image.open(self.logo_path)
        self.logo = ImageTk.PhotoImage(logo.convert("RGBA"))
        self.logo_label = tk.Label(
            self.plot_container, bd=0, bg=canvas_bg, highlightthickness=0, image=self.logo
        )
        self.logo_label.place(relx=1, rely=0, anchor="ne")
        self.logo_label.lift(self.canvas.get_tk_widget())

    def _on_configure(self, event):
        """
        Debounce window resize events to avoid excessive work.
        Only call _resize_widgets() at most once every 200ms.
        """
        if self._resize_job:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(200, self._resize_widgets)

    def _resize_widgets(self):
        """
        Rescale fonts and redraw axis text when the window is resized.
        """
        self._resize_job = None
        self._setup_scaling()
        for w in self.control_frame.winfo_children():
            try:
                w.configure(font=self.ui_style)
            except:
                pass
        # Update plot fonts
        for txt in [self.ax.title, self.ax.xaxis.label, self.ax.yaxis.label]:
            txt.set_fontsize(self.ui_style[1])
        for lbl in self.ax.get_xticklabels() + self.ax.get_yticklabels():
            lbl.set_fontsize(self.ui_style[1])
        self.canvas.draw()

    def _on_control_configure(self, event):
        """
        Adjust the scrollregion of the control canvas whenever the control
        frame changes size (for the scrollbar to work properly).
        """
        self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all"))

    def _browse_file(self):
        """
        Open a file dialog to select an Excel or Parquet file. If Excel file is selected, preview the first 15 rows
        in the Treeview so the user can pick the correct header row. If Parquet file is selected hide Treeview and populate column selectors
        """
        # Ask user for file path to file for plotting
        path = filedialog.askopenfilename(filetypes=[("Excel/Parquet files", "*.xlsx *.xls *.parquet")])
        if not path:
            return
        self.file_lbl.configure(text=path)
        # Reset state & hide header preview
        self.header_row = None
        self.time_col = None
        self.pressure_cols = []
        self.hdr_lbl.grid_remove()
        self.preview.grid_remove()
        self.time_cb.config(state="disabled", values=[])
        self.time_cb.set("")
        self.p_list.delete(0, "end")
        self.zones = []
        ext = os.path.splitext(path)[1].lower()
        if ext == ".parquet":
            try:
                df0 = pd.read_parquet(path)
                cols = list(df0.columns)
                self.time_cb.config(values=cols, state="readonly")
                for c in cols:
                    self.p_list.insert("end", c)
            except Exception as e:
                tkmsg.showerror("Error", f"Could not load Parquet:\n{e}")
        else:
            try:
                df0 = pd.read_excel(path, nrows=15, header=None)
                # Show header-selection widgets
                self.hdr_lbl.grid()
                self.preview.grid()
                cols = [f"C{c}" for c in range(df0.shape[1])]
                self.tree.config(columns=cols)
                for c in cols:
                    self.tree.heading(c, text=c)
                    self.tree.column(c, width=80, stretch=False)
                self.tree.delete(*self.tree.get_children())
                for idx, row in df0.iterrows():
                    self.tree.insert("", "end", iid=str(idx), values=list(row))
            except Exception as e:
                tkmsg.showerror("Error", f"Cannot read file:\n{e}")

    def _on_header_select(self, event):
        """
        Called when the user selects a header row in the Treeview.
        Re-read the Excel file with that header row to populate column names
        in the time dropdown and the pressure listbox.
        """
        sel = self.tree.selection()
        if not sel:
            return
        self.header_row = int(sel[0])
        self.hdr_lbl.configure(text=f"Header row: {self.header_row + 1}")
        path = self.file_lbl.cget("text")
        
        try:
            df_headers = pd.read_excel(path, header=self.header_row, nrows=3)
        except Exception as e:
            tkmsg.showerror("Error", f"Cannot read with header row {self.header_row + 1}:\n{e}")
            return

        cols = list(df_headers.columns)
        self.time_cb.config(values=cols, state="readonly")
        self.time_col = None
        self.p_list.delete(0, "end")
        for c in cols:
            self.p_list.insert("end", c)

    def _load_data_thread(self):
        """
        Spawn a background thread to load Excel data, prompt for date,
        then process and plot.
        """
        if self.time_col is None or not self.p_list.curselection():
            tkmsg.showwarning("Incomplete", "Select header, time, and pressure columns.")
            return

        self._disable_controls()
        self.collected_date_event.clear()

        # Start data reading in background
        threading.Thread(target=self._process_data, daemon=True).start()
        threading.Thread(target=self._play_loading_gif, daemon=True).start()

        # Prompt for test date (YYYY-MM-DD)
        date_str = simpledialog.askstring("Test Date", "Enter date (YYYY-MM-DD):")
        if not date_str:
            self._enable_controls()
            self.finished_loading_event.set()
            self.bad_date_event.set()
            self.collected_date_event.set()
            return

        try:
            self.test_date = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            tkmsg.showerror("Bad Date", "Date must be in format YYYY-MM-DD.")
            self._enable_controls()
            self.finished_loading_event.set()
            self.bad_date_event.set()
            self.collected_date_event.set()
            return

        self.pressure_cols = [self.p_list.get(i) for i in self.p_list.curselection()]
        self.collected_date_event.set()

    def _process_data(self):
        """
        Background worker: read the Excel file, convert time column into elapsed seconds
        (or use numeric elapsed directly), then signal when data is ready.
        """
        self.loading = True
        path = self.file_lbl.cget("text")

        if os.path.splitext(self.file_lbl.cget("text"))[-1].lower() == ".parquet":
            try:
                self.df = pd.read_parquet(path)
            except Exception:
                tkmsg.showwarning("Incomplete", "Data failed to load, cancelling.")
                return
        else:
            try:
                self.df = pd.read_excel(path, header=self.header_row)
            except Exception:
                tkmsg.showwarning("Incomplete", "Data failed to load, cancelling.")
                return

        # Wait for date input
        self.collected_date_event.wait()
        self.collected_date_event.clear()
        if self.bad_date_event.is_set():
            self.bad_date_event.clear()
            return

        if self.elapsed_mode.get():
            # Use numeric elapsed directly
            self.df[self.time_col] = pd.to_numeric(self.df[self.time_col], errors="coerce")
            self.elapsed_col = self.time_col
        else:
            # Parse absolute time: combine test_date + time strings
            self.df["ParsedTime"] = pd.to_datetime(
                self.test_date.strftime("%Y-%m-%d") + " " + self.df[self.time_col].astype(str),
                errors="coerce",
            )
            for col in self.pressure_cols:
                self.df[col] = pd.to_numeric(self.df[col], errors="coerce")
            self.df.dropna(subset=["ParsedTime"], inplace=True)
            self.elapsed_col = "Elapsed"
            self.df[self.elapsed_col] = (
                self.df["ParsedTime"] - self.df["ParsedTime"].iloc[0]
            ).dt.total_seconds()
        

        self.finished_loading_event.set()
        self.after(0, self._on_data_ready)

    def _on_data_ready(self):
        """
        Called on the main thread once _process_data finishes.
        Enable controls, set up rectangle selector, and draw initial plot.
        """
        self._enable_controls()
        self.zones = []
        self._enable_selector()
        self._redraw()

    def _enable_selector(self):
        """
        Activate the RectangleSelector so the user can draw flat zones on the plot.
        """
        if self.rs:
            self.rs.set_active(False)
            self.rs.disconnect_events()
        self.rs = RectangleSelector(
            self.ax,
            self._on_select,
            useblit=True,
            button=[1],  # Left mouse button
            minspanx=5,
            minspany=5,
            spancoords="data",
            interactive=True,
            props=dict(facecolor="red", alpha=0.3, edgecolor="black", linewidth=1),
        )
        self.rs.set_active(True)

    def _on_select(self, e1, e2):
        """
        Callback for RectangleSelector: when the user draws a rectangle (x1→x2),
        if the span exceeds min size, highlight it and label with an index.
        """
        x1, x2 = sorted([e1.xdata, e2.xdata])
        if None in (x1, x2) or x2 - x1 < self.min_var.get():
            return

        patch = self.ax.axvspan(x1, x2, color="red", alpha=0.3)
        idx = len(self.zones) + 1
        y_max = max(self.df[c].max() for c in self.pressure_cols)
        label = self.ax.text(
            (x1 + x2) / 2, y_max, str(idx), ha="center", bbox=dict(fc="yellow")
        )
        self.zones.append({"start": x1, "end": x2, "patch": patch, "label": label})
        self.canvas.draw()

    def _on_click(self, event):
        """
        If the user right-clicks (button=3) inside a drawn zone, remove that zone
        and renumber remaining ones.
        """
        if event.button != 3 or event.inaxes != self.ax:
            return
        x = event.xdata
        for i, z in enumerate(self.zones):
            if z["start"] <= x <= z["end"]:
                z["patch"].remove()
                z["label"].remove()
                self.zones.pop(i)
                break
        # Renumber labels
        for idx, z in enumerate(self.zones, 1):
            z["label"].set_text(str(idx))
            z["label"].set_x((z["start"] + z["end"]) / 2)
        self.canvas.draw()

    def _redraw(self):
        """
        Clear and redraw the entire pressure-vs-time plot: replot all pressure columns,
        then re-draw saved zones (patch + label).
        """
        self.ax.clear()
        for c in self.pressure_cols:
            self.ax.plot(self.df[self.elapsed_col], self.df[c], label=c)
        # Redraw each saved zone (if any)
        for i, z in enumerate(self.zones, 1):
            z["patch"] = self.ax.axvspan(z["start"], z["end"], color="red", alpha=0.3)
            z["label"] = self.ax.text(
                (z["start"] + z["end"]) / 2,
                max(self.df[c].max() for c in self.pressure_cols),
                str(i),
                ha="center",
                bbox=dict(fc="yellow"),
            )
        self.ax.set_xlabel("Elapsed Time [s]")
        self.ax.legend()
        self.ax.grid(True)
        self.canvas.draw()

    def _confirm(self):
        """
        When the user clicks "Confirm Zones", show a summary dialog listing each zone's
        start/end. If confirmed, open a new window per zone with time-domain and FFT plots.
        """
        if not self.zones:
            tkmsg.showwarning("No zones", "Please draw zones first.")
            return

        msgs = [f"Zone {i}: {z['start']:.2f}-{z['end']:.2f}" for i, z in enumerate(self.zones, 1)]
        if not tkmsg.askokcancel("Confirm Zones", "\n".join(msgs)):
            return

        for i, z in enumerate(self.zones, 1):
            start, end = z["start"], z["end"]
            zone_df = self.df[(self.df[self.elapsed_col] >= start) & (self.df[self.elapsed_col] <= end)].copy()
            if zone_df.empty:
                tkmsg.showerror("Zone Error", f"Zone {i} is empty.")
                continue

            # Create separate window for zone analysis
            win = tk.Toplevel(self)
            win.title(f"Zone {i} Analysis")
            win.geometry("700x900")

            fig = plt.Figure(figsize=(6, 8), dpi=100)
            ax_time = fig.add_subplot(211)
            ax_fft = fig.add_subplot(212)

            # Time-domain plot
            for col in self.pressure_cols:
                ax_time.plot(zone_df[self.elapsed_col], zone_df[col], label=col)
            ax_time.set_title(f"Zone {i} Time Series: {start:.2f}s to {end:.2f}s")
            ax_time.set_xlabel("Elapsed Time [s]")
            ax_time.set_ylabel("Pressure")
            ax_time.legend()
            ax_time.grid(True)

            # FFT plot (DC removed, scaled)
            dt = np.mean(np.diff(zone_df[self.elapsed_col].values))
            for col in self.pressure_cols:
                data = zone_df[col].values
                data = data - np.mean(data)
                N = len(data)
                freqs = np.fft.rfftfreq(N, d=dt)
                fft_vals = np.abs(np.fft.rfft(data))
                fft_vals *= 2 / N
                ax_fft.plot(freqs, fft_vals, label=col)
            ax_fft.set_title(f"Zone {i} FFT (DC Removed)")
            ax_fft.set_xlabel("Frequency [Hz]")
            ax_fft.set_ylabel("Amplitude")
            ax_fft.legend()
            ax_fft.grid(True)

            fig.tight_layout()

            # Embed figure in Tk window
            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            toolbar = NavigationToolbar2Tk(canvas, win)
            toolbar.update()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Overlay logo in the new window
            logo_widget = canvas.get_tk_widget()
            canvas_bg = logo_widget.cget("bg")
            logo_label = tk.Label(win, bd=0, bg=canvas_bg, highlightthickness=0, image=self.logo)
            logo_label.place(relx=0, rely=0, anchor="nw")
            logo_label.lift(canvas.get_tk_widget())

    def _save_analysis(self):
        """
        Save either:
          • Parquet (raw DataFrame including elapsed column + any other columns),
          • or PDF report (zones summary + plots).
        """
        if self.df is None:
            tkmsg.showwarning("No Data", "Please load data before saving.")
            return
        
        if self.save_data_mode.get():
            ext = ".parquet"
        else:
            ext = ".pdf"

        save_path = filedialog.asksaveasfile(mode='wb', title="Save As...", defaultextension=ext)

        if ext == ".parquet":
            # ——— Parquet save ———
            if not save_path:
                return
            try:
                # Save the entire DataFrame (with elapsed_col but without original "time" column) to Parquet
                self.df.drop(columns=[self.time_col])
                self.df.to_parquet(save_path, index=False)
                tkmsg.showinfo("Saved", f"DataFrame saved to {save_path.name}")
            except Exception as e:
                tkmsg.showerror("Save Error", f"Could not save Parquet:\n{e}")
        else:
            # ——— PDF report ———
            if not save_path:
                return
            try:
                with PdfPages(save_path) as pdf:
                    # Page 1: summary with logo
                    fig_sum = plt.figure(figsize=(8.27, 11.69))
                    fig_sum.clf()
                    logo = Image.open(self.logo_path)
                    logo_arr = np.array(logo)
                    ax_logo = fig_sum.add_axes([0.75, 0.85, 0.2, 0.1], anchor="NE", zorder=1)
                    ax_logo.imshow(logo_arr)
                    ax_logo.axis("off")

                    original = self.file_lbl.cget("text")
                    wrapped_path = "\n".join([original[i : i + 50] for i in range(0, len(original), 50)])

                    text = []
                    text.append("Alpha Analysis Report")
                    text.append(f"Date of Test: {self.test_date.strftime('%Y-%m-%d') if self.test_date else 'N/A'}")
                    text.append("Original File:")
                    text.append(wrapped_path)
                    text.append(f"Pressure Columns: {', '.join(self.pressure_cols)}")
                    text.append("\nZone Summary:")
                    if not self.zones:
                        text.append("None")
                    else:
                        for i, z in enumerate(self.zones, 1):
                            text.append(f"Zone {i}: {z['start']:.2f}s to {z['end']:.2f}s")
                    fig_sum.text(0.05, 0.5, "\n".join(text), ha="left", va="center", fontsize=10)
                    pdf.savefig(fig_sum)
                    plt.close(fig_sum)

                    # Page 2: overall plot with zones
                    fig_all = plt.figure(figsize=(8.27, 11.69))
                    ax_all = fig_all.add_subplot(111)
                    for c in self.pressure_cols:
                        ax_all.plot(self.df[self.elapsed_col], self.df[c], label=c)
                    for i, z in enumerate(self.zones, 1):
                        ax_all.axvspan(z["start"], z["end"], color="red", alpha=0.3)
                        ax_all.text(
                            (z["start"] + z["end"]) / 2,
                            max(self.df[c].max() for c in self.pressure_cols) * 0.95,
                            str(i),
                            ha="center",
                            va="top",
                            bbox=dict(fc="yellow"),
                        )
                    ax_all.set_title("Overall Time Plot")
                    ax_all.set_xlabel("Elapsed Time [s]")
                    ax_all.set_ylabel("Pressure")
                    ax_all.legend()
                    ax_all.grid(True)
                    pdf.savefig(fig_all)
                    plt.close(fig_all)

                    # Pages per zone
                    for i, z in enumerate(self.zones, 1):
                        start, end = z["start"], z["end"]
                        zone_df = self.df[(self.df[self.elapsed_col] >= start) & (self.df[self.elapsed_col] <= end)].copy()
                        if zone_df.empty:
                            continue
                        fig_zone = plt.figure(figsize=(8.27, 11.69))
                        ax_time = fig_zone.add_subplot(211)
                        ax_fft = fig_zone.add_subplot(212)

                        for col in self.pressure_cols:
                            ax_time.plot(zone_df[self.elapsed_col], zone_df[col], label=col)
                        ax_time.set_title(f"Zone {i} Time Series: {start:.2f}s to {end:.2f}s")
                        ax_time.set_xlabel("Elapsed Time [s]")
                        ax_time.set_ylabel("Pressure")
                        ax_time.legend()
                        ax_time.grid(True)

                        dt = np.mean(np.diff(zone_df[self.elapsed_col].values))
                        for col in self.pressure_cols:
                            data = zone_df[col].values
                            data = data - np.mean(data)
                            N = len(data)
                            freqs = np.fft.rfftfreq(N, d=dt)
                            fft_vals = np.abs(np.fft.rfft(data))
                            fft_vals *= 2 / N
                            ax_fft.plot(freqs, fft_vals, label=col)
                        ax_fft.set_title(f"Zone {i} FFT (DC Removed)")
                        ax_fft.set_xlabel("Frequency [Hz]")
                        ax_fft.set_ylabel("Amplitude")
                        ax_fft.legend()
                        ax_fft.grid(True)

                        fig_zone.tight_layout()
                        pdf.savefig(fig_zone)
                        plt.close(fig_zone)

                tkmsg.showinfo("Saved", f"Analysis saved to {save_path}")
            except Exception as e:
                tkmsg.showerror("Save Error", f"An error occurred while saving: {e}")

    def _export_zones(self):
        """
        Export each drawn zone into its own Parquet file.
        Prompts for a folder, then writes zone_1.parquet, zone_2.parquet, etc.
        """
        if self.df is None or not self.zones:
            tkmsg.showwarning("Nothing to Export", "Load data and draw zones first.")
            return

        # Ask the user for a directory in which to save each zone
        folder = filedialog.askdirectory(title="Select Folder to Save Zones")
        if not folder:
            return  # user canceled

        count = 0
        for i, z in enumerate(self.zones, start=1):
            start, end = z["start"], z["end"]
            # Slice out the DataFrame rows where elapsed_col ∈ [start, end]
            zone_df = self.df[
                (self.df[self.elapsed_col] >= start) &
                (self.df[self.elapsed_col] <= end)
            ].copy()

            if zone_df.empty:
                continue

            # Construct a filename: zone_1.parquet, zone_2.parquet, ...
            out_path = os.path.join(folder, f"zone_{i}.parquet")
            try:
                zone_df.to_parquet(out_path, index=False)
                count += 1
            except Exception as e:
                tkmsg.showerror(
                    "Export Error",
                    f"Could not save zone {i} to Parquet:\n{e}"
                )
                return

        tkmsg.showinfo(
            "Export Complete",
            f"Successfully exported {count} zone(s) to:\n{folder}"
        )

    def _get_loading_frames(self):
        """
        Load all frames from the loading GIF into a list for animation.
        """
        gif = Image.open(self.loading_gif_path)
        frames = []
        try:
            while True:
                frame = ImageTk.PhotoImage(gif.copy().convert("RGBA"))
                frames.append(frame)
                gif.seek(len(frames))
        except EOFError:
            pass
        return frames

    def _next_frame(self):
        """
        Advance to the next GIF frame. If still loading, schedule another update.
        Otherwise hide the loading label.
        """
        if self.loading_gif_frames:
            self.current_frame = (self.current_frame + 1) % len(self.loading_gif_frames)
            self.loading_label.config(image=self.loading_gif_frames[self.current_frame])
            if not self.finished_loading_event.is_set():
                self.loading_label.after(33, self._next_frame)
            else:
                self.finished_loading_event.clear()
                self.loading_label.place_forget()

    def _play_loading_gif(self):
        """
        Show the loading GIF in the center of the plotting canvas while data is loading.
        """
        if not self.loading_gif_frames:
            self.loading_gif_frames = self._get_loading_frames()
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lift(self.canvas.get_tk_widget())
        self.loading_label.config(image=self.loading_gif_frames[0])
        self._next_frame()

    def _disable_controls(self):
        """
        Disable all interactive controls during long operations (e.g., loading).
        """
        self.browse_btn.configure(state="disabled")
        self.load_btn.configure(state="disabled")
        self.confirm_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self.save_data_switch.configure(state="disabled")
        self.hdr_lbl.configure(state="disabled")
        self.tree.state(["disabled"])
        try:
            self.tree.unbind("<<TreeviewSelect>>")
        except Exception:
            pass
        self.elapsed_switch.configure(state="disabled")
        self.time_cb.configure(state="disabled")
        self.p_list.configure(state="disabled")
        self.min_entry.configure(state="disabled")

    def _enable_controls(self):
        """
        Re-enable all interactive controls after operations finish.
        """
        self.browse_btn.configure(state="normal")
        self.load_btn.configure(state="normal")
        self.confirm_btn.configure(state="normal")
        self.save_btn.configure(state="normal")
        self.save_data_switch.configure(state="normal")
        self.hdr_lbl.configure(state="normal")

        self.tree.state(["!disabled"])
        self.tree.bind("<<TreeviewSelect>>", self._on_header_select)
        self.elapsed_switch.configure(state="normal")
        if self.header_row is not None:
            self.time_cb.configure(state="readonly")
        else:
            self.time_cb.configure(state="disabled")
        self.p_list.configure(state="normal")
        self.min_entry.configure(state="normal")

    def _on_closing(self):
        """
        Prompt the user to confirm quitting the application.
        """
        if self._resize_job:
            self.after_cancel(self._resize_job)
            self._resize_job = None
        if tkmsg.askokcancel("Quit", "Do you really want to quit?"):
            self.quit()

    # ────────────────────────────────────────────────────────────────────────────
    # 7) UPDATE MECHANISM (full-installer behavior)
    # ────────────────────────────────────────────────────────────────────────────
    def _check_for_updates(self, autoUpdating = False):
        """
        Fetch update_info.json from remote, compare versions, and if newer,
        download and launch new installer.
        """
        try:
            from urllib.request import urlopen
        except ImportError:
            tkmsg.showerror("Update Error", "Cannot perform update check on this platform.")
            return

        try:
            with urlopen(UPDATE_INFO_URL, timeout=10) as resp:
                info = load(resp)
        except Exception as e:
            tkmsg.showerror("Update Error", f"Could not reach update server:\n{e}")
            return

        remote_version = info.get("version", "")
        download_url = info.get("download_url", "")

        def version_tuple(v):
            return tuple(int(x) for x in v.split("."))

        try:
            if version_tuple(remote_version) <= version_tuple(__version__):
                if not autoUpdating:
                    tkmsg.showinfo("Up To Date", f"You already have version {__version__}.")
                return
        except Exception:
            if remote_version == __version__:
                if not autoUpdating:
                    tkmsg.showinfo("Up To Date", f"You already have version {__version__}.")
                return

        if not tkmsg.askyesno(
            "Update Available",
            f"Version {remote_version} is available. You have {__version__}.\nDownload and install now?",
        ):
            return

        if getattr(sys, "frozen", False):
            install_dir = os.path.dirname(sys.executable)
        else:
            tkmsg.showerror("Not Installed", "This update mechanism only works in the bundled EXE.")
            return

        new_exe_path = os.path.join(install_dir, f"AlphaAnalysisApp_{remote_version}.exe")
        try:
            with urlopen(download_url, timeout=60) as resp:
                data = resp.read()
        except Exception as e:
            tkmsg.showerror("Download Error", f"Could not download new installer:\n{e}")
            return

        try:
            with open(new_exe_path, "wb") as f:
                f.write(data)
        except Exception as e:
            tkmsg.showerror("File Error", f"Could not save new installer:\n{e}")
            return

        old_exe = sys.executable
        try:
            os.startfile(f'"{new_exe_path}" --replace-old "{old_exe}"')
        except Exception:
            try:
                from subprocess import Popen
                Popen([new_exe_path, "--replace-old", old_exe])
            except Exception as e:
                tkmsg.showerror("Launch Error", f"Could not launch new installer:\n{e}")
                return

        self.quit()
        os._exit(0)


def main():
    """
    Create and run the AlphaAnalysisApp.
    """
    app = AlphaAnalysisApp()
    app.mainloop()


if __name__ == "__main__":
    main()
