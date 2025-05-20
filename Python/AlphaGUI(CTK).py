"""
Alpha Analysis GUI
Author: Ethan Predinchuk
Date: 2025-05-16
Version: 1.0

Description: This script provides a GUI for analyzing pressure data from excel 
files produced by the current "WELL WHISPERER" testing system. It allows users to select
an excel file, specify the header row, select time and pressure columns for analysis, and visualize the data.
The GUI also allows users to draw zones on the plot, which specify regions of interest for 
frequency analysis. A popup window displays the time-domain and frequency-domain plots for each selected zone.

Dependencies:
- tkinter
- customtkinter
- pandas
- numpy
- matplotlib
- threading
- datetime
- Pillow

NOTE:   This code is designed to be run in a Python environment with the required libraries installed.
        The code is structured to be modular, with separate functions for each part of the GUI and data processing.
        There are no external dependencies other than the standard Python libraries and the specified third-party libraries.
        There is no reason the code shouldn't work for other excel sheets, and should perform in full, but was designed for the "WELL WHISPERER" system.
"""

import threading
import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, filedialog, simpledialog
from tkinter import messagebox as tkmsg
import pandas as pd
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
from matplotlib.widgets import RectangleSelector
from datetime import datetime
from PIL import Image, ImageTk

# Appearance setup
ctk.set_appearance_mode('System')
ctk.set_default_color_theme('blue')

class AlphaAnalysisApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Alpha Analysis (Optimized)")
        self.geometry("1200x800")
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Debounce resize
        self._resize_job = None
        self.bind('<Configure>', self._on_configure)

        # Base dimensions for scaling
        self.base_width = 1200
        self.base_height = 800
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
        self.collected_date_event = threading.Event()
        self.bad_date_event = threading.Event()

        # Elapsed switch
        self.elapsed_mode = tk.BooleanVar(value=False)

        # Create control frame area with scroll bar
        self.control_container = ctk.CTkFrame(self)
        self.control_container.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
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
            width=int(self.winfo_width()*0.8),
        )

        self.control_canvas.configure(bg=self.control.cget("fg_color")[1])
        self.control.bind("<Configure>", self._on_control_configure)

        self.timePlot = ctk.CTkFrame(self)
        self.timePlot.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self._build_controls()
        self._build_plot()

        # Loading Gif holder
        canvas_widget = self.canvas.get_tk_widget()
        canvas_bg = canvas_widget.cget('bg')
        self.loading_gif_path = "Images/LoadingGIF.gif"
        self.loading_gif_frames = []
        self.current_frame = 0
        self.loading_label = tk.Label(self.timePlot, bd=0, bg=canvas_bg, highlightthickness=0)
        self.finished_loading_event = threading.Event()

    def _build_controls(self):
        # Browse button
        ctk.CTkLabel(self.control, text="1. Select Excel File", anchor='w', font=self.ui_style).pack(fill='x')
        self.browse_btn = ctk.CTkButton(self.control, text="Browse...", command=self._browse_file, font=self.ui_style)
        self.browse_btn.pack(fill='x', pady=5)
        self.file_lbl = ctk.CTkLabel(self.control, text="No file chosen", wraplength=280, anchor='w', font=self.ui_style)
        self.file_lbl.pack(fill='x', pady=5)

        # File preview and header row selection
        ctk.CTkLabel(self.control, text="2. Choose Header Row", anchor='w', font=self.ui_style).pack(fill='x')
        self.preview = tk.Frame(self.control, height=180)
        self.preview.pack(fill='x', pady=5)
        self.preview.pack_propagate(False)
        self.tree = ttk.Treeview(self.preview, show='headings', height=8)
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
        self.hdr_lbl.pack(fill='x', pady=5)

        # Elapsed time toggle switch
        self.elapsed_switch = ctk.CTkSwitch(self.control, text="Use Elapsed Only", variable=self.elapsed_mode, font=self.ui_style)
        self.elapsed_switch.pack(anchor='w', pady=5)

        # Dropdown time column selector
        ctk.CTkLabel(self.control, text="3. Select Columns", anchor='w', font=self.ui_style).pack(fill='x')
        ctk.CTkLabel(self.control, text="Time Column:", anchor='w', font=self.ui_style).pack(fill='x')
        self.time_cb = ttk.Combobox(self.control, state='disabled')
        self.time_cb.pack(fill='x', pady=5)
        self.time_cb.bind('<<ComboboxSelected>>', lambda e: setattr(self, 'time_col', self.time_cb.get()))

        # Multiselect pressure columns
        ctk.CTkLabel(self.control, text="Pressure Columns:", anchor='w', font=self.ui_style).pack(fill='x')
        self.p_list = tk.Listbox(self.control, selectmode='multiple', height=5, font=self.ui_style)
        self.p_list.pack(fill='x', pady=5)

        # Load and plot button
        self.load_btn = ctk.CTkButton(self.control, text="4. Load & Plot", command=self._load_data_thread, font=self.ui_style)
        self.load_btn.pack(fill='x', pady=5)
        
        # Minimum zone size box [TODO: Add autocorrect to 0 when invalid input is entered]
        ctk.CTkLabel(self.control, text="Min Zone Size (s):", anchor='w', font=self.ui_style).pack(fill='x')
        self.min_var = tk.DoubleVar(value=30.0)
        self.min_entry = ctk.CTkEntry(self.control, textvariable=self.min_var, font=self.ui_style)
        self.min_entry.pack(fill='x', pady=5)

        # Confirm zones and plot frequency responses button
        self.confirm_btn = ctk.CTkButton(self.control, text="Confirm Zones", command=self._confirm, font=self.ui_style)
        self.confirm_btn.pack(fill='x', pady=5)

    def _build_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(6,5))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.timePlot)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas, self.timePlot)
        self.rs = None
        self.canvas.mpl_connect('button_press_event', self._on_click)

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
        self.data_thread = threading.Thread(target=self._process_data, daemon=True).start()
        self.loading_thread = threading.Thread(target=self._play_loading_gif, daemon=True).start()

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
                self.loading_label.after(33, self._next_frame) # 33ms delay for GIF frames = 30 fps
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

if __name__=='__main__':
    app=AlphaAnalysisApp() 
    app.mainloop()