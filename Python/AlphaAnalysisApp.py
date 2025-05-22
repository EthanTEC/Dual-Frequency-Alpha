from threading import Thread, Event
import tkinter as tk
from zipfile import ZipFile
import customtkinter as ctk
from tkinter import ttk, filedialog, simpledialog
from tkinter import messagebox as tkmsg
import pandas as pd
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
from matplotlib.widgets import RectangleSelector
from datetime import datetime
from PIL import Image, ImageTk
from json import dump, JSONDecodeError, load
from urllib.request import URLError, HTTPError, urlopen
import os
import sys
import shutil
from textwrap import wrap
from tempfile import NamedTemporaryFile

# Application version
__version__ = "1.0.1"
# URL where the current version info is stored (should return JSON with {'version': 'x.y.z', 'url': 'http://.../AlphaAnalysisApp.py'})
UPDATE_INFO_URL = "https://raw.githubusercontent.com/EthanTEC/Dual-Frequency-Alpha/refs/heads/main/Python/update_info.json"

# Appearance setup
ctk.set_appearance_mode('dark')
ctk.set_default_color_theme('blue')

# Determine base path so that "Images/…" can be loaded both in dev and once bundled:
if getattr(sys, "frozen", False):
    # Running as a PyInstaller‐frozen executable
    BASE_PATH = sys._MEIPASS
else:
    # When running from source, __file__ is ".../Dual Frequency Alpha/Python/AlphaAnalysisApp.py"
    # We want BASE_PATH = ".../Dual Frequency Alpha", so go one level up.
    script_dir = os.path.abspath(os.path.dirname(__file__))
    BASE_PATH = os.path.abspath(os.path.join(script_dir, os.pardir))

class AlphaAnalysisApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Alpha Analysis (Optimized)")
        self.geometry("1600x900")
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Debounce resize
        self._resize_job = None
        self.bind('<Configure>', self._on_configure)

        # Base dimensions for scaling
        self.base_width = 1600
        self.base_height = 900
        self.base_font_size = 12
        self.ui_font = "Segoe UI"
        self.ui_style = (self.ui_font, self.base_font_size)

        # Style for ttk Treeview
        self.ttk_style = ttk.Style(self)
        self.ttk_style.configure('Treeview', font=(self.ui_font, self.base_font_size), rowheight=self.base_font_size*2)
        self.ttk_style.configure('Treeview.Heading', font=("Segoe UI", self.base_font_size//2, 'bold'))

        # Data placeholders
        self.df = None
        self.zones = []  # dicts: {'start','end','patch','label'}
        self.time_col = None
        self.pressure_cols = []
        self.elapsed_col = None
        self.test_date = None
        self.header_row = None
        self.collected_date_event = Event()
        self.bad_date_event = Event()

        # Elapsed switch
        self.elapsed_mode = tk.BooleanVar(value=False)

        # Save mode switch
        self.save_data_mode = tk.BooleanVar(value=False)

        # Create control frame area with scroll bar
        self.control_container = ctk.CTkFrame(self, width=250)
        self.control_container.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=2)
        self.control_container.pack_propagate(False)

        self.control_canvas = tk.Canvas(self.control_container, borderwidth=0, highlightthickness=0)
        self.control_scrollbar = ttk.Scrollbar(self.control_container, orient="vertical", command=self.control_canvas.yview)
        self.control_canvas.configure(yscrollcommand=self.control_scrollbar.set)

        self.control_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.control_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.control = ctk.CTkFrame(self.control_canvas)
        self.control_window = self.control_canvas.create_window(
            (0, 0),
            window=self.control,
            anchor="nw",
            width=int(self.winfo_width()),
        )

        self.control_canvas.configure(bg=self.control.cget("fg_color")[1])
        self.control.bind("<Configure>", self._on_control_configure)

        self.timePlot = ctk.CTkFrame(self)
        self.timePlot.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self._build_controls()
        self._build_plot()

        #Run an update check in the background on startup
        Thread(target=self._check_for_updates_background, daemon=True).start()

    def _build_controls(self):
        # Browse button
        ctk.CTkLabel(self.control, text="1. Select Excel File", anchor='w', font=self.ui_style).pack(fill='x')
        self.browse_btn = ctk.CTkButton(self.control, text="Browse...", command=self._browse_file, font=self.ui_style)
        self.browse_btn.pack(fill='x', pady=2)
        self.file_lbl = ctk.CTkLabel(self.control, text="No file chosen", wraplength=280, anchor='w', font=self.ui_style)
        self.file_lbl.pack(fill='x', pady=2)

        # File preview and header row selection
        ctk.CTkLabel(self.control, text="2. Choose Header Row", anchor='w', font=self.ui_style).pack(fill='x')
        self.preview = tk.Frame(self.control, height=180)
        self.preview.pack(fill='x', pady=2)
        self.preview.pack_propagate(False)
        self.tree = ttk.Treeview(self.preview, show='headings', height=6)
        vs = ttk.Scrollbar(self.preview, orient='vertical', command=self.tree.yview)
        hs = ttk.Scrollbar(self.preview, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscroll=vs.set, xscroll=hs.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        hs.grid(row=1, column=0, sticky='ew')
        self.preview.grid_rowconfigure(0, weight=1)
        self.preview.grid_columnconfigure(0, weight=1)
        self.tree.bind('<<TreeviewSelect>>', self._on_header_select)
        self.hdr_lbl = ctk.CTkLabel(self.control, text="Header row: None", anchor='w', font=self.ui_style)
        self.hdr_lbl.pack(fill='x', pady=2)

        # Elapsed time toggle switch
        self.elapsed_switch = ctk.CTkSwitch(self.control, text="Use Elapsed Only", variable=self.elapsed_mode, font=self.ui_style)
        self.elapsed_switch.pack(anchor='w', pady=2)

        # Dropdown time column selector
        ctk.CTkLabel(self.control, text="3. Select Columns", anchor='w', font=self.ui_style).pack(fill='x')
        ctk.CTkLabel(self.control, text="Time Column:", anchor='w', font=self.ui_style).pack(fill='x')
        self.time_cb = ttk.Combobox(self.control, state='disabled')
        self.time_cb.pack(fill='x', pady=2)
        self.time_cb.bind('<<ComboboxSelected>>', lambda e: setattr(self, 'time_col', self.time_cb.get()))

        # Multiselect pressure columns
        ctk.CTkLabel(self.control, text="Pressure Columns:", anchor='w', font=self.ui_style).pack(fill='x')
        self.p_list = tk.Listbox(self.control, selectmode='multiple', height=5, font=self.ui_style)
        self.p_list.pack(fill='x', pady=2)

        # Load and plot button
        self.load_btn = ctk.CTkButton(self.control, text="4. Load & Plot", command=self._load_data_thread, font=self.ui_style)
        self.load_btn.pack(fill='x', pady=2)
        
        # Minimum zone size box
        ctk.CTkLabel(self.control, text="Min Zone Size (s):", anchor='w', font=self.ui_style).pack(fill='x')
        self.min_var = tk.DoubleVar(value=30.0)
        self.min_entry = ctk.CTkEntry(self.control, textvariable=self.min_var, font=self.ui_style)
        self.min_entry.pack(fill='x', pady=2)

        # Confirm zones and plot frequency responses button
        self.confirm_btn = ctk.CTkButton(self.control, text="5. Confirm Zones", command=self._confirm, font=self.ui_style)
        self.confirm_btn.pack(fill='x', pady=2)

        # Save options label and toggle
        ctk.CTkLabel(self.control, text="Save Options", anchor='w', font=self.ui_style).pack(fill='x', pady=2)
        self.save_data_switch = ctk.CTkSwitch(self.control, text="Save as data", variable=self.save_data_mode, font=self.ui_style)
        self.save_data_switch.pack(anchor='w', pady=2)

        # Save analysis / data button
        self.save_btn = ctk.CTkButton(self.control, text="6. Save", command=self._save_analysis, font=self.ui_style)
        self.save_btn.pack(fill='x', pady=2)

        # Check for updates button
        self.update_btn = ctk.CTkButton(self.control, text="7. Check for Updates", command=self._check_for_updates, font=self.ui_style)
        self.update_btn.pack(fill='x', pady=2)

    def _build_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(6,5))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.timePlot)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas, self.timePlot)
        self.rs = None
        self.canvas.mpl_connect('button_press_event', self._on_click)
        
        # Loading Gif holder
        loading_widget = self.canvas.get_tk_widget()
        canvas_bg = loading_widget.cget('bg')
        self.loading_gif_path = os.path.join(BASE_PATH, "Images", "LoadingGIF.gif")
        self.loading_gif_frames = []
        self.current_frame = 0
        self.loading_label = tk.Label(self.timePlot, bd=0, bg=canvas_bg, highlightthickness=0)
        self.finished_loading_event = Event()

        # Logo display
        logo_widget = self.canvas.get_tk_widget()
        canvas_bg = logo_widget.cget('bg')
        self.logo_path = os.path.join(BASE_PATH, "Images", "TEC.jpg")
        logo = Image.open(self.logo_path)
        self.logo = ImageTk.PhotoImage(logo.convert('RGBA'))
        self.logo_label = tk.Label(self.timePlot, bd=0, bg=canvas_bg, highlightthickness=0, image=self.logo)
        self.logo_label.place(relx=1, rely=0, anchor='ne')
        self.logo_label.lift(self.canvas.get_tk_widget())

    def _on_configure(self, event):
        if self._resize_job:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(200, self._resize_widgets)

    def _resize_widgets(self):
        self._resize_job = None
        w = self.winfo_width() or self.base_width
        h = self.winfo_height() or self.base_height
        scale = min(w/self.base_width, h/self.base_height)
        new_size = max(6, min(int(self.base_font_size * scale), 20))
        self.ui_style = ("Segoe UI", new_size)
        # Update CTk widgets
        for widget in self.control.winfo_children():
            try: widget.configure(font=self.ui_style)
            except: pass
        # Update ttk Treeview
        self.ttk_style.configure('Treeview', font=(self.ui_font, self.base_font_size), rowheight=new_size*2)
        self.ttk_style.configure('Treeview.Heading', font=("Segoe UI", new_size//2, 'bold'))
        # Update listbox
        self.p_list.config(font=("Segoe UI", new_size))
        # Update plot fonts
        if hasattr(self, 'ax'):
            for txt in [self.ax.title, self.ax.xaxis.label, self.ax.yaxis.label]:
                txt.set_fontsize(new_size)
            for lbl in self.ax.get_xticklabels() + self.ax.get_yticklabels():
                lbl.set_fontsize(new_size)
            self.canvas.draw()

    def _on_control_configure(self, event):
        """
        Update the scrollregion of the canvas whenever the control frame is resized.
        """
        self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all"))

    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")])
        if not path: return
        self.file_lbl.configure(text=path)
        df0 = pd.read_excel(path, nrows=15, header=None)
        cols = [f"C{c}" for c in range(df0.shape[1])]
        self.tree.config(columns=cols)
        for c in cols:
            self.tree.heading(c, text=c); self.tree.column(c, width=80, stretch=False)
        self.tree.delete(*self.tree.get_children())
        for idx, row in df0.iterrows():
            self.tree.insert('', 'end', iid=str(idx), values=list(row))
        self.hdr_lbl.configure(text="Header row: None")
        self.time_cb.config(state='disabled'); self.p_list.delete(0, 'end')

    def _on_header_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        self.header_row = int(sel[0]); self.hdr_lbl.configure(text=f"Header row: {self.header_row+1}")
        path = self.file_lbl.cget('text')
        try:
            df_headers = pd.read_excel(path, header=self.header_row, nrows=3)
        except Exception:
            tkmsg.showerror("Error","Cannot read with header row {i+1}")
            return
        cols = list(df_headers.columns)
        self.time_cb.config(values=cols, state='readonly')
        self.time_col = None
        self.p_list.delete(0,'end')
        for c in cols:
            self.p_list.insert('end', c)

    def _load_data_thread(self):

        # Only start processing thread if you have all the selections needed to plot
        if self.header_row is None or self.time_col is None or not self.p_list.curselection():
            tkmsg.showwarning("Incomplete","Select header, time, and pressure columns.")
            return
        
        # Disable control panel while processing data
        self._disable_controls()

        # Process data and play loading animation in separate threads to avoid blocking the UI
        self.collected_date_event.clear()
        Thread(target=self._process_data, daemon=True).start()
        Thread(target=self._play_loading_gif, daemon=True).start()

        # ask date in main thread
        date_str = simpledialog.askstring("Test Date","Enter date (YYYY-MM-DD):")
        if not date_str:
            self._enable_controls()
            self.finished_loading_event.set()
            self.bad_date_event.set()
            self.collected_date_event.set()
            return
        try: 
            self.test_date = datetime.strptime(date_str,'%Y-%m-%d')
        except: 
            tkmsg.showerror("Bad Date","Date must be YYYY-MM-DD.")
            self._enable_controls()
            self.finished_loading_event.set()
            self.bad_date_event.set()
            self.collected_date_event.set()
            return
        
        # collect cols
        self.pressure_cols = [self.p_list.get(i) for i in self.p_list.curselection()]
        self.collected_date_event.set()
        
    def _process_data(self):

        # Flag you are loading data
        self.loading = True

        path = self.file_lbl.cget('text')
        parsed_data = pd.read_excel(path, header=self.header_row)

        if parsed_data is None:
            tkmsg.showwarning("Incomplete","Data Failed to load, cancelling.")
            return
        
        self.collected_date_event.wait()
        self.collected_date_event.clear()

        if self.bad_date_event.is_set():
            self.bad_date_event.clear()
            return

        self.df = parsed_data

        if self.elapsed_mode.get(): # Elapsed mode
            self.df[self.time_col] = pd.to_numeric(self.df[self.time_col], errors='coerce', downcast='float')
            self.df.dropna(subset=[self.time_col], inplace=True)
            self.elapsed_col = self.time_col

        else: # Absolute mode
            self.df['ParsedTime'] = pd.to_datetime(
                self.test_date.strftime('%Y-%m-%d') + ' ' + self.df[self.time_col].astype(str),
                errors='coerce')
            self.df.dropna(subset=['ParsedTime'], inplace=True)
            self.elapsed_col = 'Elapsed'
            self.df[self.elapsed_col] = (self.df['ParsedTime'] - self.df['ParsedTime'].iloc[0]).dt.total_seconds()
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
            spancoords='data',
            interactive=True,
            props=dict(facecolor='red', alpha=0.3, edgecolor='black', linewidth=1)
        )
        self.rs.set_active(True)

    def _on_select(self, e1, e2):
        x1,x2=sorted([e1.xdata,e2.xdata])
        if None in (x1,x2) or x2-x1 < self.min_var.get(): return
        patch=self.ax.axvspan(x1,x2,color='red',alpha=0.3)
        idx=len(self.zones)+1
        y_max=max(self.df[c].max() for c in self.pressure_cols)
        label=self.ax.text((x1+x2)/2,y_max,str(idx),ha='center',bbox=dict(fc='yellow'))
        self.zones.append({'start':x1,'end':x2,'patch':patch,'label':label})
        self.canvas.draw()

    def _on_click(self,event):
        if event.button!=3 or event.inaxes!=self.ax: return
        x=event.xdata
        for i,z in enumerate(self.zones):
            if z['start']<=x<=z['end']:
                z['patch'].remove(); z['label'].remove(); self.zones.pop(i); break
        for idx,z in enumerate(self.zones,1):
            z['label'].set_text(str(idx)); z['label'].set_x((z['start']+z['end'])/2)
        self.canvas.draw()

    def _redraw(self):
        self.ax.clear()
        for c in self.pressure_cols:
            self.ax.plot(self.df[self.elapsed_col],self.df[c],label=c)
        for i,z in enumerate(self.zones,1):
            z['patch']=self.ax.axvspan(z['start'],z['end'],color='red',alpha=0.3)
            z['label']=self.ax.text((z['start']+z['end'])/2,max(self.df[c].max() for c in self.pressure_cols),str(i),ha='center',bbox=dict(fc='yellow'))
        self.ax.set_xlabel('Elapsed Time [s]')
        self.ax.legend(); self.ax.grid(True)
        self.canvas.draw()

    def _confirm(self):
        if not self.zones: tkmsg.showwarning("No zones","Please draw zones first."); return

        msgs=[f"Zone {i}: {z['start']:.2f}-{z['end']:.2f}" for i,z in enumerate(self.zones,1)]
        if not tkmsg.askokcancel("Confirm Zones","\n".join(msgs)): return

        for i, z in enumerate(self.zones, 1):
            start, end = z['start'], z['end']
            zone_df = self.df[(self.df[self.elapsed_col] >= start) & (self.df[self.elapsed_col] <= end)].copy()
            if zone_df.empty:
                tkmsg.showerror("Zone Error", f"Zone {i} is empty.")
                continue

            # Create Toplevel window
            win = tk.Toplevel(self)
            win.title(f"Zone {i} Analysis")
            win.geometry("700x900")

            # Create matplotlib Figure
            fig = plt.Figure(figsize=(6, 8), dpi=100)
            ax_time = fig.add_subplot(211)
            ax_fft = fig.add_subplot(212)

            # Time-domain plot
            for col in self.pressure_cols:
                ax_time.plot(zone_df[self.elapsed_col], zone_df[col], label=col)
            ax_time.set_title(f"Zone {i} Time Series: {start:.2f}s to {end:.2f}s")
            ax_time.set_xlabel('Elapsed Time [s]')
            ax_time.set_ylabel('Pressure')
            ax_time.legend()
            ax_time.grid(True)

            # FFT plot (DC removed and scaled)
            dt = np.mean(np.diff(zone_df[self.elapsed_col].values))
            for col in self.pressure_cols:
                data = zone_df[col].values
                data = data - np.mean(data)  # Remove DC
                N = len(data)
                freqs = np.fft.rfftfreq(N, d=dt)
                fft_vals = np.abs(np.fft.rfft(data))
                fft_vals *= 2 / N
                ax_fft.plot(freqs, fft_vals, label=col)
            ax_fft.set_title(f"Zone {i} of {self.pressure_cols} FFT (DC Removed)")
            ax_fft.set_xlabel('Frequency [Hz]')
            ax_fft.set_ylabel('Amplitude')
            ax_fft.legend()
            ax_fft.grid(True)

            fig.tight_layout()

            # Embed in Tkinter
            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            toolbar = NavigationToolbar2Tk(canvas, win)
            toolbar.update()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Logo display
            logo_widget = canvas.get_tk_widget()
            canvas_bg = logo_widget.cget('bg')
            logo_label = tk.Label(win, bd=0, bg=canvas_bg, highlightthickness=0, image=self.logo)
            logo_label.place(relx=0, rely=0, anchor='nw')
            logo_label.lift(canvas.get_tk_widget())

    def _save_analysis(self):
        # Ensure data and zones exist
        if self.df is None:
            tkmsg.showwarning("No Data", "Please load data before saving.")
            return
        if not self.zones:
            tkmsg.showwarning("No Zones", "Please draw and confirm zones before saving.")
            return

        # Ask user for folder and filename
        folder = filedialog.askdirectory(title="Select Save Folder")
        if not folder:
            return
        filename = simpledialog.askstring("Filename", "Enter filename (without extension):")
        if not filename:
            return

        # If saving as data (JSON)
        if self.save_data_mode.get():
            save_path = f"{folder}/{filename}.json"
            try:
                # Convert DataFrame to JSON
                df_json = self.df.to_json(orient='split')
                data_to_save = {
                    'dataframe': df_json,
                    'time_col': self.time_col,
                    'pressure_cols': self.pressure_cols,
                    'elapsed_col': self.elapsed_col,
                    'test_date': self.test_date.strftime('%Y-%m-%d') if self.test_date else None,
                    'header_row': self.header_row,
                    'zones': [
                        {'start': z['start'], 'end': z['end']} for z in self.zones
                    ],
                    'original_file': self.file_lbl.cget('text')
                }
                with open(save_path, 'w') as f:
                    dump(data_to_save, f)
                tkmsg.showinfo("Saved", f"Data saved to {save_path}")
            except Exception as e:
                tkmsg.showerror("Save Error", f"An error occurred while saving data: {e}")
        else:
            # Save as report (PDF)
            save_path = f"{folder}/{filename}.pdf"
            try:
                with PdfPages(save_path) as pdf:
                    # First page: summary with date, original file path, and logo
                    fig_sum = plt.figure(figsize=(8.27, 11.69))  # A4 size
                    fig_sum.clf()
                    # Add logo in upper-right corner
                    logo = Image.open(self.logo_path)
                    logo_arr = np.array(logo)
                    ax_logo = fig_sum.add_axes([0.75, 0.85, 0.2, 0.1], anchor='NE', zorder=1)
                    ax_logo.imshow(logo_arr)
                    ax_logo.axis('off')

                    # Prepare wrapped file path
                    original = self.file_lbl.cget('text')
                    wrapped_path = '\n'.join(wrap(original, width=50))

                    text = []
                    text.append(f"Alpha Analysis Report")
                    text.append(f"Date of Test: {self.test_date.strftime('%Y-%m-%d') if self.test_date else 'N/A'}")
                    text.append(f"Original File:")
                    text.append(wrapped_path)
                    text.append(f"Pressure Columns: {', '.join(self.pressure_cols)}")
                    text.append("\nZone Summary:")
                    for i, z in enumerate(self.zones, 1):
                        text.append(f"Zone {i}: {z['start']:.2f}s to {z['end']:.2f}s")
                    fig_sum.text(0.05, 0.5, '\n'.join(text), ha='left', va='center', fontsize=10)
                    pdf.savefig(fig_sum)
                    plt.close(fig_sum)

                    # Overall time plot with highlighted zones
                    fig_all = plt.figure(figsize=(8.27, 11.69))
                    ax_all = fig_all.add_subplot(111)
                    for c in self.pressure_cols:
                        ax_all.plot(self.df[self.elapsed_col], self.df[c], label=c)
                    # Highlight zones
                    for i, z in enumerate(self.zones, 1):
                        ax_all.axvspan(z['start'], z['end'], color='red', alpha=0.3)
                        ax_all.text((z['start']+z['end'])/2, max(self.df[c].max() for c in self.pressure_cols)*0.95, str(i), ha='center', va='top', bbox=dict(fc='yellow'))
                    ax_all.set_title('Overall Time Plot')
                    ax_all.set_xlabel('Elapsed Time [s]')
                    ax_all.set_ylabel('Pressure')
                    ax_all.legend()
                    ax_all.grid(True)
                    pdf.savefig(fig_all)
                    plt.close(fig_all)

                    # Zone-specific plots
                    for i, z in enumerate(self.zones, 1):
                        start, end = z['start'], z['end']
                        zone_df = self.df[(self.df[self.elapsed_col] >= start) & (self.df[self.elapsed_col] <= end)].copy()
                        if zone_df.empty:
                            continue
                        fig_zone = plt.figure(figsize=(8.27, 11.69))
                        ax_time = fig_zone.add_subplot(211)
                        ax_fft = fig_zone.add_subplot(212)

                        # Time plot
                        for col in self.pressure_cols:
                            ax_time.plot(zone_df[self.elapsed_col], zone_df[col], label=col)
                        ax_time.set_title(f"Zone {i} Time Series: {start:.2f}s to {end:.2f}s")
                        ax_time.set_xlabel('Elapsed Time [s]')
                        ax_time.set_ylabel('Pressure')
                        ax_time.legend()
                        ax_time.grid(True)

                        # FFT
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
                        ax_fft.set_xlabel('Frequency [Hz]')
                        ax_fft.set_ylabel('Amplitude')
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
                frame = ImageTk.PhotoImage(gif.copy().convert('RGBA'))
                frames.append(frame)
                gif.seek(len(frames))  # move to next frame
        except EOFError:
            pass  # end of sequence
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
        self.loading_label.place(relx=0.5, rely=0.5, anchor='center')
        self.loading_label.lift(self.canvas.get_tk_widget())
        self.loading_label.config(image=self.loading_gif_frames[0])
        self._next_frame()

    def _disable_controls(self):
        """
        Disable (grey-out) all interactive widgets in the left-pane,
        preventing user changes while loading/plotting is in progress.
        """
        # Disable each button
        self.browse_btn.configure(state='disabled')
        self.load_btn.configure(state='disabled')
        self.confirm_btn.configure(state='disabled')
        self.save_btn.configure(state='disabled')
        self.save_data_switch.configure(state='disabled')

        # Disable the Treeview (header selection) via its state method
        self.tree.state(['disabled'])
        try:
            self.tree.unbind('<<TreeviewSelect>>')
        except Exception:
            pass

        # Disable the "Use Elapsed Only" switch
        self.elapsed_switch.configure(state='disabled')

        # Disable ComboBox & Listbox
        self.time_cb.configure(state='disabled')
        self.p_list.configure(state='disabled')

        # Disable the min zone size entry
        self.min_entry.configure(state='disabled')

    def _enable_controls(self):
        """
        Re-enable all previously disabled widgets once loading is complete.
        """
        # Re‐enable buttons
        self.browse_btn.configure(state='normal')
        self.load_btn.configure(state='normal')
        self.confirm_btn.configure(state='normal')
        self.save_btn.configure(state='normal')
        self.save_data_switch.configure(state='normal')

        # Re‐enable Treeview and re‐bind its event via its state method
        self.tree.state(['!disabled'])
        self.tree.bind('<<TreeviewSelect>>', self._on_header_select)

        # Re‐enable the switch
        self.elapsed_switch.configure(state='normal')

        # Re‐enable ComboBox (only readonly if header was chosen) & Listbox
        if self.header_row is not None:
            self.time_cb.configure(state='readonly')
        else:
            self.time_cb.configure(state='disabled')
        self.p_list.configure(state='normal')

        # Re‐enable the min zone size entry
        self.min_entry.configure(state='normal')

    def _on_closing(self):
        if self._resize_job:
            self.after_cancel(self._resize_job)
            self._resize_job = None
        if tkmsg.askokcancel("Quit", "Do you really want to quit?"):
            self.quit()

    def _check_for_updates(self):
        """
        Called when the user clicks "Check for Updates".
        Handles both full build and patch-based updates.
        """
        # 1) Fetch update_info.json
        try:
            with urlopen(UPDATE_INFO_URL, timeout=10) as resp:
                info = load(resp)
        except Exception as e:
            tkmsg.showerror("Update Error", f"Could not reach update server:\n{e}")
            return

        remote_version = info.get("version", "")
        full_url       = info.get("full_url", "")
        patch_url      = info.get("patch_url", "")
        is_exe         = info.get("is_exe", False)

        def version_tuple(v): return tuple(int(x) for x in v.split("."))

        # Compare versions
        try:
            if version_tuple(remote_version) <= version_tuple(__version__):
                tkmsg.showinfo("Up To Date", f"You already have version {__version__}.")
                return
        except:
            # fallback on simple string compare
            if remote_version == __version__:
                tkmsg.showinfo("Up To Date", f"You already have version {__version__}.")
                return

        # New version found
        choice = tkmsg.askyesnocancel(
            "Update Available",
            f"Version {remote_version} is available.  You have {__version__}.\n\n"
            "Click YES to download a small patch (recommended) if available.\n"
            "Click NO to download the full installer ZIP instead.\n"
            "Click CANCEL to skip."
        )
        if choice is None:
            return  # user clicked "Cancel"
        if choice is True and patch_url:
            # USER CHOSE PATCH
            download_url = patch_url
            update_type = "patch"
        else:
            # USER CHOSE FULL (or no patch_url available)
            download_url = full_url
            update_type = "full"

        if not download_url:
            tkmsg.showerror("No Download URL", "No download URL is configured for this update.")
            return

        # Download the ZIP into a temp file
        try:
            with urlopen(download_url, timeout=60) as resp:
                data = resp.read()
        except Exception as e:
            tkmsg.showerror("Download Error", f"Could not download update:\n{e}")
            return

        # Write to a temporary ZIP on disk
        tmp_zip = NamedTemporaryFile(delete=False, suffix=".zip")
        tmp_zip.write(data)
        tmp_zip.flush()
        tmp_zip.close()
        zip_path = tmp_zip.name

        # Determine installation directory:
        if getattr(sys, "frozen", False):
            # When frozen, the EXE sits inside, say, C:/Program Files/AlphaAnalysisApp/AlphaAnalysisApp.exe
            exe_path = sys.executable
            install_dir = os.path.abspath(os.path.dirname(exe_path))
        else:
            # In dev, assume the user is running the onedir under dist/AlphaAnalysisApp/
            # i.e. Python/AlphaAnalysisApp.py was never installed, so we ask them to pick folder:
            install_dir = filedialog.askdirectory(title="Select Save Folder")
            if not install_dir or not os.path.isdir(install_dir):
                tkmsg.showerror("Invalid Folder", "Please select a valid installation folder.")
                os.remove(zip_path)
                return
            exe_path = os.path.join(install_dir, "AlphaAnalysisApp.exe")
            if not os.path.isfile(exe_path):
                tkmsg.showerror("Not Installed", "Could not find AlphaAnalysisApp.exe in that folder.")
                os.remove(zip_path)
                return
        try:
            # Extract zip into install_dir, overwriting files
            with ZipFile(zip_path, 'r') as z:
                z.extractall(install_dir)
        except Exception as e:
            tkmsg.showerror("Patch Error", f"Could not apply update:\n{e}")
            os.remove(zip_path)
            return
        finally:
            os.remove(zip_path)

        tkmsg.showinfo("Updated", f"Application updated to version {remote_version}.\nRestarting now...")

        # Relaunch the exe
        if getattr(sys, "frozen", False):
            os.execl(exe_path, exe_path, *sys.argv)
        else:
            # In dev mode, launch the EXE inside the onedir:
            os.execl(exe_path, exe_path)

    def _check_for_updates_background(self):
        """
        Called on startup in a background thread. Just notifies if a newer version exists.
        """
        try:
            with urlopen(UPDATE_INFO_URL, timeout=10) as resp:
                info = load(resp)
        except:
            return  # ignore offline or invalid JSON

        remote_version = info.get("version", "")
        try:
            if tuple(int(x) for x in remote_version.split(".")) <= tuple(int(x) for x in __version__.split(".")):
                return
        except:
            if remote_version == __version__:
                return

        def notify():
            if tkmsg.askyesno(
                "Update Available",
                f"A new version {remote_version} is available. You have {__version__}.\n"
                "Click Yes to download."
            ):
                self._check_for_updates()
        self.after(1000, notify)

if __name__=='__main__':
    app=AlphaAnalysisApp() 
    app.mainloop()
