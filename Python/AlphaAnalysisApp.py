# AlphaAnalysisApp.py
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
from zipfile import ZipFile
from json import dump, load, JSONDecodeError
from urllib.request import urlopen, URLError, HTTPError
from tempfile import NamedTemporaryFile
import shutil

# ──────────────────────────────────────────────────────────────────────────────
# 1) VERSION AND UPDATE_INFO_URL
# ──────────────────────────────────────────────────────────────────────────────
__version__ = "1.0.0"
UPDATE_INFO_URL = "https://raw.githubusercontent.com/EthanTEC/Dual-Frequency-Alpha/main/Python/update_info.json"

# ──────────────────────────────────────────────────────────────────────────────
# 2) HANDLE --replace-old ARGUMENT (delete the old EXE)
# ──────────────────────────────────────────────────────────────────────────────
def try_delete_old_exe():
    """
    If the script was launched with "--replace-old <old_path>", wait a moment
    and then delete <old_path>.  This allows the new EXE to overwrite/delete
    the previous version.
    """
    # Example args: ["AlphaAnalysisApp.exe", "--replace-old", "C:/Apps/AlphaAnalysisApp.exe"]
    args = sys.argv
    if "--replace-old" in args:
        idx = args.index("--replace-old")
        if idx + 1 < len(args):
            old_path = args[idx + 1]
            # Allow a short delay so that Windows releases locks on old EXE
            time.sleep(1.0)
            try:
                if os.path.isfile(old_path):
                    os.remove(old_path)
            except Exception:
                pass
            # After deleting, remove the flags from sys.argv so they don't confuse the rest
            del sys.argv[idx: idx + 2]

# Always run this at startup, before anything else
try_delete_old_exe()

# ──────────────────────────────────────────────────────────────────────────────
# 3) DETERMINE BASE_PATH FOR LOADING IMAGES (works in-dev and after PyInstaller)
# ──────────────────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_PATH = sys._MEIPASS
else:
    script_dir = os.path.abspath(os.path.dirname(__file__))
    BASE_PATH = os.path.abspath(os.path.join(script_dir, os.pardir))

# ──────────────────────────────────────────────────────────────────────────────
# 4) APPLICATION CLASS
# ──────────────────────────────────────────────────────────────────────────────
class AlphaAnalysisApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Alpha Analysis (Optimized)")
        self.geometry("1600x900")
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Debounce resize
        self._resize_job = None
        self.bind("<Configure>", self._on_configure)

        # Base dimensions for scaling
        self.base_width = 1600
        self.base_height = 900
        self.base_font_size = 12
        self.ui_font = "Segoe UI"
        self.ui_style = (self.ui_font, self.base_font_size)

        # Style for ttk Treeview
        self.ttk_style = ttk.Style(self)
        self.ttk_style.configure(
            "Treeview", font=(self.ui_font, self.base_font_size), rowheight=self.base_font_size * 2
        )
        self.ttk_style.configure(
            "Treeview.Heading", font=("Segoe UI", self.base_font_size // 2, "bold")
        )

        # Data placeholders
        self.df = None
        self.zones = []
        self.time_col = None
        self.pressure_cols = []
        self.elapsed_col = None
        self.test_date = None
        self.header_row = None
        self.collected_date_event = threading.Event()
        self.bad_date_event = threading.Event()

        # Elapsed switch
        self.elapsed_mode = tk.BooleanVar(value=False)
        # Save mode switch
        self.save_data_mode = tk.BooleanVar(value=False)

        # Control frame with scrollbar
        self.control_container = ctk.CTkFrame(self, width=250)
        self.control_container.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=2)
        self.control_container.pack_propagate(False)

        self.control_canvas = tk.Canvas(self.control_container, borderwidth=0, highlightthickness=0)
        self.control_scrollbar = ttk.Scrollbar(
            self.control_container, orient="vertical", command=self.control_canvas.yview
        )
        self.control_canvas.configure(yscrollcommand=self.control_scrollbar.set)
        self.control_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.control_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.control = ctk.CTkFrame(self.control_canvas)
        self.control_window = self.control_canvas.create_window(
            (0, 0), window=self.control, anchor="nw", width=int(self.winfo_width())
        )
        self.control_canvas.configure(bg=self.control.cget("fg_color")[1])
        self.control.bind("<Configure>", self._on_control_configure)

        self.timePlot = ctk.CTkFrame(self)
        self.timePlot.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self._build_controls()
        self._build_plot()

    def _build_controls(self):
        # 1) Browse button
        ctk.CTkLabel(self.control, text="1. Select Excel File", anchor="w", font=self.ui_style).pack(fill="x")
        self.browse_btn = ctk.CTkButton(self.control, text="Browse...", command=self._browse_file, font=self.ui_style)
        self.browse_btn.pack(fill="x", pady=2)
        self.file_lbl = ctk.CTkLabel(self.control, text="No file chosen", wraplength=280, anchor="w", font=self.ui_style)
        self.file_lbl.pack(fill="x", pady=2)

        # 2) Header row selector
        ctk.CTkLabel(self.control, text="2. Choose Header Row", anchor="w", font=self.ui_style).pack(fill="x")
        self.preview = tk.Frame(self.control, height=180)
        self.preview.pack(fill="x", pady=2)
        self.preview.pack_propagate(False)
        self.tree = ttk.Treeview(self.preview, show="headings", height=6)
        vs = ttk.Scrollbar(self.preview, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(self.preview, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vs.set, xscroll=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.preview.grid_rowconfigure(0, weight=1)
        self.preview.grid_columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_header_select)
        self.hdr_lbl = ctk.CTkLabel(self.control, text="Header row: None", anchor="w", font=self.ui_style)
        self.hdr_lbl.pack(fill="x", pady=2)

        # Elapsed toggle
        self.elapsed_switch = ctk.CTkSwitch(self.control, text="Use Elapsed Only", variable=self.elapsed_mode, font=self.ui_style)
        self.elapsed_switch.pack(anchor="w", pady=2)

        # 3) Time & pressure column selectors
        ctk.CTkLabel(self.control, text="3. Select Columns", anchor="w", font=self.ui_style).pack(fill="x")
        ctk.CTkLabel(self.control, text="Time Column:", anchor="w", font=self.ui_style).pack(fill="x")
        self.time_cb = ttk.Combobox(self.control, state="disabled")
        self.time_cb.pack(fill="x", pady=2)
        self.time_cb.bind("<<ComboboxSelected>>", lambda e: setattr(self, "time_col", self.time_cb.get()))
        ctk.CTkLabel(self.control, text="Pressure Columns:", anchor="w", font=self.ui_style).pack(fill="x")
        self.p_list = tk.Listbox(self.control, selectmode="multiple", height=5, font=self.ui_style)
        self.p_list.pack(fill="x", pady=2)

        # 4) Load & Plot button
        self.load_btn = ctk.CTkButton(self.control, text="4. Load & Plot", command=self._load_data_thread, font=self.ui_style)
        self.load_btn.pack(fill="x", pady=2)

        # Min zone size
        ctk.CTkLabel(self.control, text="Min Zone Size (s):", anchor="w", font=self.ui_style).pack(fill="x")
        self.min_var = tk.DoubleVar(value=30.0)
        self.min_entry = ctk.CTkEntry(self.control, textvariable=self.min_var, font=self.ui_style)
        self.min_entry.pack(fill="x", pady=2)

        # 5) Confirm zones
        self.confirm_btn = ctk.CTkButton(self.control, text="5. Confirm Zones", command=self._confirm, font=self.ui_style)
        self.confirm_btn.pack(fill="x", pady=2)

        # Save options
        ctk.CTkLabel(self.control, text="Save Options", anchor="w", font=self.ui_style).pack(fill="x", pady=2)
        self.save_data_switch = ctk.CTkSwitch(self.control, text="Save as data", variable=self.save_data_mode, font=self.ui_style)
        self.save_data_switch.pack(anchor="w", pady=2)

        # 6) Save analysis/data
        self.save_btn = ctk.CTkButton(self.control, text="6. Save", command=self._save_analysis, font=self.ui_style)
        self.save_btn.pack(fill="x", pady=2)

        # 7) Check for Updates (full installer)
        self.update_btn = ctk.CTkButton(self.control, text="7. Check for Updates", command=self._check_for_updates, font=self.ui_style)
        self.update_btn.pack(fill="x", pady=2)

    def _build_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(6, 5))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.timePlot)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas, self.timePlot)
        self.rs = None
        self.canvas.mpl_connect("button_press_event", self._on_click)

        # Loading GIF
        loading_widget = self.canvas.get_tk_widget()
        canvas_bg = loading_widget.cget("bg")
        self.loading_gif_path = os.path.join(BASE_PATH, "Images", "LoadingGIF.gif")
        self.loading_gif_frames = []
        self.current_frame = 0
        self.loading_label = tk.Label(self.timePlot, bd=0, bg=canvas_bg, highlightthickness=0)
        self.finished_loading_event = threading.Event()

        # Logo
        logo_widget = self.canvas.get_tk_widget()
        canvas_bg = logo_widget.cget("bg")
        self.logo_path = os.path.join(BASE_PATH, "Images", "TEC.jpg")
        logo = Image.open(self.logo_path)
        self.logo = ImageTk.PhotoImage(logo.convert("RGBA"))
        self.logo_label = tk.Label(self.timePlot, bd=0, bg=canvas_bg, highlightthickness=0, image=self.logo)
        self.logo_label.place(relx=1, rely=0, anchor="ne")
        self.logo_label.lift(self.canvas.get_tk_widget())

    def _on_configure(self, event):
        if self._resize_job:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(200, self._resize_widgets)

    def _resize_widgets(self):
        self._resize_job = None
        w = self.winfo_width() or self.base_width
        h = self.winfo_height() or self.base_height
        scale = min(w / self.base_width, h / self.base_height)
        new_size = max(6, min(int(self.base_font_size * scale), 20))
        self.ui_style = ("Segoe UI", new_size)
        for widget in self.control.winfo_children():
            try:
                widget.configure(font=self.ui_style)
            except:
                pass
        self.ttk_style.configure("Treeview", font=(self.ui_font, self.base_font_size), rowheight=new_size * 2)
        self.ttk_style.configure("Treeview.Heading", font=("Segoe UI", new_size // 2, "bold"))
        self.p_list.config(font=("Segoe UI", new_size))
        if hasattr(self, "ax"):
            for txt in [self.ax.title, self.ax.xaxis.label, self.ax.yaxis.label]:
                txt.set_fontsize(new_size)
            for lbl in self.ax.get_xticklabels() + self.ax.get_yticklabels():
                lbl.set_fontsize(new_size)
            self.canvas.draw()

    def _on_control_configure(self, event):
        self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all"))

    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        self.file_lbl.configure(text=path)
        df0 = pd.read_excel(path, nrows=15, header=None)
        cols = [f"C{c}" for c in range(df0.shape[1])]
        self.tree.config(columns=cols)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=80, stretch=False)
        self.tree.delete(*self.tree.get_children())
        for idx, row in df0.iterrows():
            self.tree.insert("", "end", iid=str(idx), values=list(row))
        self.hdr_lbl.configure(text="Header row: None")
        self.time_cb.config(state="disabled")
        self.p_list.delete(0, "end")

    def _on_header_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        self.header_row = int(sel[0])
        self.hdr_lbl.configure(text=f"Header row: {self.header_row + 1}")
        path = self.file_lbl.cget("text")
        try:
            df_headers = pd.read_excel(path, header=self.header_row, nrows=3)
        except Exception:
            tkmsg.showerror("Error", f"Cannot read with header row {self.header_row + 1}")
            return
        cols = list(df_headers.columns)
        self.time_cb.config(values=cols, state="readonly")
        self.time_col = None
        self.p_list.delete(0, "end")
        for c in cols:
            self.p_list.insert("end", c)

    def _load_data_thread(self):
        if self.header_row is None or self.time_col is None or not self.p_list.curselection():
            tkmsg.showwarning("Incomplete", "Select header, time, and pressure columns.")
            return
        self._disable_controls()
        self.collected_date_event.clear()
        threading.Thread(target=self._process_data, daemon=True).start()
        threading.Thread(target=self._play_loading_gif, daemon=True).start()

        date_str = simpledialog.askstring("Test Date", "Enter date (YYYY-MM-DD):")
        if not date_str:
            self._enable_controls()
            self.finished_loading_event.set()
            self.bad_date_event.set()
            self.collected_date_event.set()
            return
        try:
            self.test_date = datetime.strptime(date_str, "%Y-%m-%d")
        except:
            tkmsg.showerror("Bad Date", "Date must be YYYY-MM-DD.")
            self._enable_controls()
            self.finished_loading_event.set()
            self.bad_date_event.set()
            self.collected_date_event.set()
            return

        self.pressure_cols = [self.p_list.get(i) for i in self.p_list.curselection()]
        self.collected_date_event.set()

    def _process_data(self):
        self.loading = True
        path = self.file_lbl.cget("text")
        parsed_data = pd.read_excel(path, header=self.header_row)
        if parsed_data is None:
            tkmsg.showwarning("Incomplete", "Data failed to load, cancelling.")
            return
        self.collected_date_event.wait()
        self.collected_date_event.clear()
        if self.bad_date_event.is_set():
            self.bad_date_event.clear()
            return

        self.df = parsed_data
        if self.elapsed_mode.get():  # Elapsed mode
            self.df[self.time_col] = pd.to_numeric(self.df[self.time_col], errors="coerce", downcast="float")
            self.df.dropna(subset=[self.time_col], inplace=True)
            self.elapsed_col = self.time_col
        else:  # Absolute mode
            self.df["ParsedTime"] = pd.to_datetime(
                self.test_date.strftime("%Y-%m-%d") + " " + self.df[self.time_col].astype(str), errors="coerce"
            )
            self.df.dropna(subset=["ParsedTime"], inplace=True)
            self.elapsed_col = "Elapsed"
            self.df[self.elapsed_col] = (self.df["ParsedTime"] - self.df["ParsedTime"].iloc[0]).dt.total_seconds()

        self.finished_loading_event.set()
        self.after(0, self._on_data_ready)

    def _on_data_ready(self):
        self._enable_controls()
        self.zones = []
        self._enable_selector()
        self._redraw()
        self.rs.set_active(True)

    def _enable_selector(self):
        if self.rs:
            self.rs.set_active(False)
            self.rs.disconnect_events()
        self.rs = RectangleSelector(
            self.ax,
            self._on_select,
            useblit=True,
            button=[1],
            minspanx=5,
            minspany=5,
            spancoords="data",
            interactive=True,
            props=dict(facecolor="red", alpha=0.3, edgecolor="black", linewidth=1),
        )
        self.rs.set_active(True)

    def _on_select(self, e1, e2):
        x1, x2 = sorted([e1.xdata, e2.xdata])
        if None in (x1, x2) or x2 - x1 < self.min_var.get():
            return
        patch = self.ax.axvspan(x1, x2, color="red", alpha=0.3)
        idx = len(self.zones) + 1
        y_max = max(self.df[c].max() for c in self.pressure_cols)
        label = self.ax.text((x1 + x2) / 2, y_max, str(idx), ha="center", bbox=dict(fc="yellow"))
        self.zones.append({"start": x1, "end": x2, "patch": patch, "label": label})
        self.canvas.draw()

    def _on_click(self, event):
        if event.button != 3 or event.inaxes != self.ax:
            return
        x = event.xdata
        for i, z in enumerate(self.zones):
            if z["start"] <= x <= z["end"]:
                z["patch"].remove()
                z["label"].remove()
                self.zones.pop(i)
                break
        for idx, z in enumerate(self.zones, 1):
            z["label"].set_text(str(idx))
            z["label"].set_x((z["start"] + z["end"]) / 2)
        self.canvas.draw()

    def _redraw(self):
        self.ax.clear()
        for c in self.pressure_cols:
            self.ax.plot(self.df[self.elapsed_col], self.df[c], label=c)
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

            # FFT plot (DC removed and scaled)
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

            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            toolbar = NavigationToolbar2Tk(canvas, win)
            toolbar.update()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            logo_widget = canvas.get_tk_widget()
            canvas_bg = logo_widget.cget("bg")
            logo_label = tk.Label(win, bd=0, bg=canvas_bg, highlightthickness=0, image=self.logo)
            logo_label.place(relx=0, rely=0, anchor="nw")
            logo_label.lift(canvas.get_tk_widget())

    def _save_analysis(self):
        if self.df is None:
            tkmsg.showwarning("No Data", "Please load data before saving.")
            return
        if not self.zones:
            tkmsg.showwarning("No Zones", "Please draw and confirm zones before saving.")
            return

        folder = filedialog.askdirectory(title="Select Save Folder")
        if not folder:
            return
        filename = simpledialog.askstring("Filename", "Enter filename (without extension):")
        if not filename:
            return

        if self.save_data_mode.get():
            save_path = f"{folder}/{filename}.json"
            try:
                df_json = self.df.to_json(orient="split")
                data_to_save = {
                    "dataframe": df_json,
                    "time_col": self.time_col,
                    "pressure_cols": self.pressure_cols,
                    "elapsed_col": self.elapsed_col,
                    "test_date": self.test_date.strftime("%Y-%m-%d") if self.test_date else None,
                    "header_row": self.header_row,
                    "zones": [{"start": z["start"], "end": z["end"]} for z in self.zones],
                    "original_file": self.file_lbl.cget("text"),
                }
                with open(save_path, "w") as f:
                    dump(data_to_save, f)
                tkmsg.showinfo("Saved", f"Data saved to {save_path}")
            except Exception as e:
                tkmsg.showerror("Save Error", f"An error occurred while saving data: {e}")
        else:
            save_path = f"{folder}/{filename}.pdf"
            try:
                with PdfPages(save_path) as pdf:
                    fig_sum = plt.figure(figsize=(8.27, 11.69))
                    fig_sum.clf()
                    logo = Image.open(self.logo_path)
                    logo_arr = np.array(logo)
                    ax_logo = fig_sum.add_axes([0.75, 0.85, 0.2, 0.1], anchor="NE", zorder=1)
                    ax_logo.imshow(logo_arr)
                    ax_logo.axis("off")

                    original = self.file_lbl.cget("text")
                    wrapped_path = "\n".join(wrap(original, width=50))

                    text = []
                    text.append("Alpha Analysis Report")
                    text.append(f"Date of Test: {self.test_date.strftime('%Y-%m-%d') if self.test_date else 'N/A'}")
                    text.append("Original File:")
                    text.append(wrapped_path)
                    text.append(f"Pressure Columns: {', '.join(self.pressure_cols)}")
                    text.append("\nZone Summary:")
                    for i, z in enumerate(self.zones, 1):
                        text.append(f"Zone {i}: {z['start']:.2f}s to {z['end']:.2f}s")
                    fig_sum.text(0.05, 0.5, "\n".join(text), ha="left", va="center", fontsize=10)
                    pdf.savefig(fig_sum)
                    plt.close(fig_sum)

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

    def _get_loading_frames(self):
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
        if self.loading_gif_frames:
            self.current_frame = (self.current_frame + 1) % len(self.loading_gif_frames)
            self.loading_label.config(image=self.loading_gif_frames[self.current_frame])
            if not self.finished_loading_event.is_set():
                self.loading_label.after(33, self._next_frame)
            else:
                self.finished_loading_event.clear()
                self.loading_label.place_forget()

    def _play_loading_gif(self):
        if not self.loading_gif_frames:
            self.loading_gif_frames = self._get_loading_frames()
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label.lift(self.canvas.get_tk_widget())
        self.loading_label.config(image=self.loading_gif_frames[0])
        self._next_frame()

    def _disable_controls(self):
        self.browse_btn.configure(state="disabled")
        self.load_btn.configure(state="disabled")
        self.confirm_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self.save_data_switch.configure(state="disabled")
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
        self.browse_btn.configure(state="normal")
        self.load_btn.configure(state="normal")
        self.confirm_btn.configure(state="normal")
        self.save_btn.configure(state="normal")
        self.save_data_switch.configure(state="normal")
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
        if self._resize_job:
            self.after_cancel(self._resize_job)
            self._resize_job = None
        if tkmsg.askokcancel("Quit", "Do you really want to quit?"):
            self.quit()

    # ──────────────────────────────────────────────────────────────────────────────
    # 5) FULL‐INSTALLER UPDATE LOGIC (_check_for_updates)
    # ──────────────────────────────────────────────────────────────────────────────
    def _check_for_updates(self):
        """
        Called when the user clicks "Check for Updates".
        Downloads a small JSON (update_info.json), compares versions,
        then—if a newer version exists—downloads the full new EXE,
        saves it as AlphaAnalysisApp_new.exe, launches it with
        --replace-old <current_exe_path>, and exits this process.
        """
        # 1) Fetch update_info.json
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

        # 2) Compare versions
        try:
            if version_tuple(remote_version) <= version_tuple(__version__):
                tkmsg.showinfo("Up To Date", f"You already have version {__version__}.")
                return
        except:
            if remote_version == __version__:
                tkmsg.showinfo("Up To Date", f"You already have version {__version__}.")
                return

        # 3) Prompt user to download
        if not tkmsg.askyesno(
            "Update Available",
            f"Version {remote_version} is available. You have {__version__}.\n"
            "Download and install now?",
        ):
            return

        # 4) Download the new EXE into a temp file
        #    But we want to place it next to the current EXE in the same folder,
        #    named "AlphaAnalysisApp_new.exe" so we can launch it from there.
        if getattr(sys, "frozen", False):
            install_dir = os.path.dirname(sys.executable)
        else:
            tkmsg.showerror("Not Installed", "This update mechanism only works in the bundled EXE.")
            return

        new_exe_path = os.path.join(install_dir, "AlphaAnalysisApp_new.exe")

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

        # 5) Launch the new EXE with --replace-old <old_exe_path>
        old_exe = sys.executable  # path to current running EXE
        try:
            # On Windows, os.startfile is simplest
            os.startfile(f'"{new_exe_path}" --replace-old "{old_exe}"')
        except Exception:
            # Fallback to subprocess
            try:
                import subprocess
                subprocess.Popen([new_exe_path, "--replace-old", old_exe])
            except Exception as e:
                tkmsg.showerror("Launch Error", f"Could not launch new installer:\n{e}")
                return

        # 6) Exit current application immediately
        self.quit()
        os._exit(0)

# ──────────────────────────────────────────────────────────────────────────────
# 6) ENTRY POINT
# ──────────────────────────────────────────────────────────────────────────────
def main():
    app = AlphaAnalysisApp()
    app.mainloop()


if __name__ == "__main__":
    main()
