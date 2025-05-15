import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
from matplotlib.widgets import RectangleSelector
import tkinter.font as tkfont
from datetime import datetime

# Initialize CustomTkinter Appearance
ctk.set_appearance_mode('System')
ctk.set_default_color_theme('blue')

class AlphaAnalysisApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Alpha Analysis Application")
        self.geometry("1600x900")

        # Data attributes
        self.df = None
        self.zones = [] # list of dicts: {'start', 'end', 'patch', 'label'}
        self.header_row = None
        self.time_col = None
        self.pressure_cols = []
        self.elapsed_col = None
        self.test_date = None  # store for export

        # Elapsed-mode toggle
        self.elapsed_mode = tk.BooleanVar(value=False)  # False=Absolute, True=Elapsed

        # Layout frames
        self.control = ctk.CTkFrame(self)
        self.control.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        self.control.pack_propagate(False)
        self.plot_f = ctk.CTkFrame(self)
        self.plot_f.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self._build_controls()
        self._build_plot()

    def _build_controls(self):
        ctk.CTkLabel(self.control, text="1. Select Excel File", anchor='w').pack(fill='x', pady=(0,5))
        ctk.CTkButton(self.control, text="Browse...", command=self._select_file).pack(fill='x', pady=(0,10))
        self.file_lbl = ctk.CTkLabel(self.control, text="No file chosen", wraplength=280, anchor='w')
        self.file_lbl.pack(fill='x', pady=(0,10))

        ctk.CTkLabel(self.control, text="2. Preview & Choose Header Row", anchor='w').pack(fill='x')
        self.preview = ctk.CTkFrame(self.control)
        self.preview.configure(height=180)
        self.preview.pack(fill='x', pady=(5,10))
        self.preview.pack_propagate(False)
        self.tree = ttk.Treeview(self.preview, show='headings', height=8, selectmode='browse')
        vs = ttk.Scrollbar(self.preview, orient=tk.VERTICAL, command=self.tree.yview)
        hs = ttk.Scrollbar(self.preview, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=vs.set, xscroll=hs.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        hs.grid(row=1, column=0, sticky='ew')
        self.preview.grid_rowconfigure(0, weight=1)
        self.preview.grid_columnconfigure(0, weight=1)
        self.tree.bind('<<TreeviewSelect>>', self._on_header_select)
        self.hdr_lbl = ctk.CTkLabel(self.control, text="Header row: None", anchor='w')
        self.hdr_lbl.pack(fill='x', pady=(0,10))

        # Elapsed toggle switch
        ctk.CTkLabel(self.control, text="Elapsed Time Mode:", anchor='w').pack(fill='x')
        ctk.CTkSwitch(self.control,
                      text="Use Elapsed Only",
                      variable=self.elapsed_mode,
                      onvalue=True,
                      offvalue=False).pack(anchor='w', pady=(0,10))

        ctk.CTkLabel(self.control, text="3. Select Time & Pressure Columns", anchor='w').pack(fill='x')
        ctk.CTkLabel(self.control, text="Time Column:", anchor='w').pack(fill='x')
        self.time_cb = ttk.Combobox(self.control, state='disabled')
        self.time_cb.pack(fill='x')
        self.time_cb.bind('<<ComboboxSelected>>', self._on_time_select)
        ctk.CTkLabel(self.control, text="Pressure Columns:", anchor='w').pack(fill='x', pady=(10,0))
        self.p_list = tk.Listbox(self.control, selectmode='multiple', height=5)
        self.p_list.pack(fill='x')
        self.p_list.bind('<<ListboxSelect>>', self._on_pressure_select)

        ctk.CTkButton(self.control, text="4. Load & Plot", command=self._load_data).pack(fill='x', pady=(10,10))
        ctk.CTkLabel(self.control, text="Min Zone Size (s):", anchor='w').pack(fill='x')
        self.min_var = tk.DoubleVar(value=30.0)
        ctk.CTkEntry(self.control, textvariable=self.min_var).pack(fill='x')

        ctk.CTkButton(self.control, text="Confirm Zones", command=self._confirm).pack(fill='x', pady=(10,0))

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")])
        if not path: return
        self.file_lbl.configure(text=path)
        df = pd.read_excel(path, nrows=15, header=None)
        cols = [f"C{c}" for c in range(df.shape[1])]
        self.tree.config(columns=cols)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=80, stretch=False)
        self.tree.delete(*self.tree.get_children())
        for idx, row in df.iterrows():
            self.tree.insert('', tk.END, iid=str(idx), values=list(row))
        self.hdr_lbl.configure(text="Header row: None")
        self.time_cb.config(state='disabled')
        self.p_list.delete(0, tk.END)
        self.header_row = None
        self.time_col = None
        self.pressure_cols = []

    def _on_header_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        self.header_row = int(sel[0])
        self.hdr_lbl.configure(text=f"Header row: {self.header_row+1}")
        path = self.file_lbl.cget('text')
        try:
            df = pd.read_excel(path, header=self.header_row)
        except Exception:
            messagebox.showerror("File Error", "Failed to read the file with selected header row.")
            return
        self.df = df
        cols = list(df.columns)
        self.time_cb.config(values=cols, state='readonly')
        self.time_cb.set("")
        self.time_col = None
        self.p_list.delete(0, tk.END)
        for c in cols:
            self.p_list.insert(tk.END, c)
        self.pressure_cols = []


    def _build_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(6,5))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_f)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas, self.plot_f)
        self.rs = None
        self.canvas.mpl_connect('button_press_event', self._on_click)

    def _on_resize(self, event):
        w = self.winfo_width() or self.base_width
        self.control.config(width=w//4)
        new_size = max(6, int(self.base_font_size * w / self.base_width))
        self.ui_font.configure(size=new_size)

    def _on_click(self, event):
        #Right-click to clear zones only after data is loaded
        if event.button != 3 or event.inaxes != self.ax or self.df is None:
            return
        x = event.xdata
        for zi, z in enumerate(self.zones):
            if z['start'] <= x <= z['end']:
                z['patch'].remove()
                z['label'].remove()
                self.zones.pop(zi)
                break
        # Renumber zones
        for idx, z in enumerate(self.zones, 1):
            z['label'].set_text(str(idx))
            mid = (z['start'] + z['end']) / 2
            z['label'].set_x(mid)
        self.canvas.draw()



    def _on_time_select(self, event):
        self.time_col = self.time_cb.get()

    def _on_pressure_select(self, event):
        self.pressure_cols = [self.p_list.get(i) for i in self.p_list.curselection()]

    def _load_data(self):
        if self.df is None or not self.time_col or not self.pressure_cols:
            messagebox.showwarning("Incomplete", "Ensure header, time, and pressure columns chosen.")
            return

        date_str = simpledialog.askstring("Test Date", "Enter date (YYYY-MM-DD):")
        if not date_str:
            return
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
        except Exception:
            messagebox.showerror("Bad Date", "Date must be YYYY-MM-DD.")
            return

        try:
            self.df['ParsedTime'] = pd.to_datetime(
                date_str + ' ' + self.df[self.time_col].astype(str),
                format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
        except Exception:
            messagebox.showerror("Parse Error", "Time column could not be parsed.")
            return

        self.df.dropna(subset=['ParsedTime'], inplace=True)
        if self.df.empty:
            messagebox.showerror("Parse Error", "No valid times found.")
            return

        self.elapsed_col = 'Elapsed'
        self.df[self.elapsed_col] = (self.df['ParsedTime'] - self.df['ParsedTime'].iloc[0]).dt.total_seconds()
        self.zones = []
        self._enable_selector()
        self._redraw()

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
        if self.df is None or e1.xdata is None or e2.xdata is None:
            return
        x1, x2 = sorted([e1.xdata, e2.xdata])
        if x2 - x1 < self.min_var.get():
            return
        patch = self.ax.axvspan(x1, x2, color='red', alpha=0.3)
        idx = len(self.zones) + 1
        y_max = max(self.df[col].max() for col in self.pressure_cols)
        label = self.ax.text((x1+x2)/2, y_max, str(idx), ha='center', va='top', bbox=dict(fc='yellow'))
        self.zones.append({'start':x1, 'end':x2, 'patch':patch, 'label':label})
        self.canvas.draw()

    def _redraw(self):
        self.ax.clear()
        if self.df is None:
            self.canvas.draw()
            return
        for col in self.pressure_cols:
            self.ax.plot(self.df[self.elapsed_col], self.df[col], label=col)
        for idx, z in enumerate(self.zones, 1):
            patch = self.ax.axvspan(z['start'], z['end'], color='red', alpha=0.3)
            y_max = max(self.df[col].max() for col in self.pressure_cols)
            label = self.ax.text((z['start']+z['end'])/2, y_max, str(idx), ha='center', va='top', bbox=dict(fc='yellow'))
            z['patch'], z['label'] = patch, label
        self.ax.legend()
        self.ax.set_xlabel('Elapsed Time [s]')
        self.ax.grid(True)
        self.canvas.draw()

    def _confirm(self):

        # Ensure there are zones to confirm
        if not self.zones:
            messagebox.showwarning("No zones selected", "Please create at least one zone before confirming.")
            return
        
        # Prepare zone messages
        msgs = []
        for idx, z in enumerate(self.zones, 1):
            msgs.append(f"Zone {idx}: {z['start']:.2f}s to {z['end']:.2f}s")
        if not messagebox.askokcancel("Confirm Zones", "\n".join(msgs)): # Stop if user cancels
            return

        # For each selected zone, create a Toplevel window with embedded plots
        for i, z in enumerate(self.zones, 1):
            start, end = z['start'], z['end']
            zone_df = self.df[(self.df[self.elapsed_col] >= start) & (self.df[self.elapsed_col] <= end)].copy()
            if zone_df.empty:
                messagebox.showerror("Zone Error", f"Zone {i} is empty.")
                continue

            # Create Toplevel window
            win = tk.Toplevel(self)
            win.title(f"Zone {i} Analysis")
            win.geometry("700x800")

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

    def _on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.destroy()
            self.quit()

if __name__ == '__main__':
    app = AlphaAnalysisApp()
    app.protocol("WM_DELETE_WINDOW", app._on_closing)
    app.mainloop()
