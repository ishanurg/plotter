import sys, io, os, re, json, time, math
import numpy as np
import pandas as pd
from scipy import stats, optimize, fft
from scipy.fft import fft, fftfreq
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
from tkinter.scrolledtext import ScrolledText

# ---------- tksheet (professional spreadsheet) ----------
try:
    from tksheet import Sheet
    SHEET_AVAILABLE = True
except ImportError:
    SHEET_AVAILABLE = False
    print("tksheet not installed. Run: pip install tksheet")
    print("Using fallback table (limited spreadsheet)")

# ---------- matplotlib for graphing ----------
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

# Set matplotlib style for a clean, modern look
plt.style.use('seaborn-v0_8-darkgrid')
matplotlib.rcParams['font.size'] = 10

# ============================================================
# Main Application Class
# ============================================================
class DataGraphApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OriginLab Pro – Spreadsheet & Graphing")
        self.geometry("1400x900")
        self.minsize(800, 600)

        # ---------- Style configuration ----------
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure(bg='#f0f0f0')
        self.style.configure('TNotebook', background='#f0f0f0')
        self.style.configure('TNotebook.Tab', padding=[12, 4], font=('Segoe UI', 10, 'bold'))
        self.style.map('TNotebook.Tab', background=[('selected', '#ffffff')])

        # ---------- Data Store ----------
        self.df = pd.DataFrame(columns=['A','B','C'])
        self.undo_stack = []

        # ---------- Build UI ----------
        self.create_menu()
        self.create_notebook()

        # Bind keyboard shortcuts
        self.bind_shortcuts()

        # Status bar
        self.status = ttk.Label(self, text="Ready", relief=tk.SUNKEN, anchor=tk.W, padding=4)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # Load sample data to show something
        self.load_sample_data()

    # ============================================================
    # Menu Bar
    # ============================================================
    def create_menu(self):
        menubar = tk.Menu(self)

        # File
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Workbook", command=self.new_workbook, accelerator="Ctrl+N")
        file_menu.add_command(label="Open CSV...", command=self.open_csv, accelerator="Ctrl+O")
        file_menu.add_command(label="Save as CSV...", command=self.save_csv, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        # Edit
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo", command=self.undo, accelerator="Ctrl+Z")
        edit_menu.add_command(label="Redo", command=self.redo, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="Insert Row", command=self.insert_row)
        edit_menu.add_command(label="Delete Row", command=self.delete_row)
        edit_menu.add_command(label="Insert Column", command=self.insert_column)
        edit_menu.add_command(label="Delete Column", command=self.delete_column)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        # Data
        data_menu = tk.Menu(menubar, tearoff=0)
        data_menu.add_command(label="Sort Ascending", command=lambda: self.sort_column(True))
        data_menu.add_command(label="Sort Descending", command=lambda: self.sort_column(False))
        data_menu.add_separator()
        data_menu.add_command(label="Fill Down", command=self.fill_down)
        data_menu.add_command(label="Fill Right", command=self.fill_right)
        menubar.add_cascade(label="Data", menu=data_menu)

        # Analysis
        analysis_menu = tk.Menu(menubar, tearoff=0)
        analysis_menu.add_command(label="Descriptive Statistics", command=self.show_descriptive_stats)
        analysis_menu.add_command(label="Curve Fitting", command=self.curve_fitting_dialog)
        analysis_menu.add_command(label="FFT", command=self.show_fft)
        analysis_menu.add_command(label="Histogram", command=self.show_histogram)
        menubar.add_cascade(label="Analysis", menu=analysis_menu)

        # Graph
        graph_menu = tk.Menu(menubar, tearoff=0)
        graph_menu.add_command(label="Plot Selected Columns", command=self.plot_selected, accelerator="Ctrl+P")
        graph_menu.add_command(label="Clear Graph", command=self.clear_graph)
        menubar.add_cascade(label="Graph", menu=graph_menu)

        self.config(menu=menubar)

    # ============================================================
    # Notebook (tabs)
    # ============================================================
    def create_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ---------- Spreadsheet tab ----------
        self.sheet_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet_frame, text="📊 Spreadsheet")
        self.build_spreadsheet()

        # ---------- Graph tab ----------
        self.graph_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.graph_frame, text="📈 Graph")
        self.build_graph_tab()

    # ============================================================
    # Spreadsheet Construction
    # ============================================================
    def build_spreadsheet(self):
        if SHEET_AVAILABLE:
            self.sheet = Sheet(self.sheet_frame,
                               data=[],
                               headers=[],
                               show_row_index=True,
                               show_top_left_corner=False,
                               empty_vertical=0,
                               empty_horizontal=0,
                               theme="light blue")
            self.sheet.enable_bindings(("single_select",
                                        "row_select",
                                        "column_width_resize",
                                        "arrowkeys",
                                        "right_click_popup_menu",
                                        "rc_select",
                                        "copy",
                                        "paste",
                                        "cut",
                                        "delete",
                                        "undo",
                                        "edit_cell"))
            self.sheet.pack(fill=tk.BOTH, expand=True)
        else:
            # Fallback: simple Text widget (very basic)
            self.sheet_frame.grid_rowconfigure(0, weight=1)
            self.sheet_frame.grid_columnconfigure(0, weight=1)
            self.text_sheet = ScrolledText(self.sheet_frame, wrap=tk.NONE)
            self.text_sheet.grid(row=0, column=0, sticky="nsew")
            self.sheet = None  # we'll handle differently

    def update_sheet_display(self):
        """Transfer pandas DataFrame to tksheet."""
        if SHEET_AVAILABLE and self.sheet:
            data = self.df.values.tolist()
            headers = list(self.df.columns)
            self.sheet.set_sheet_data(data)
            self.sheet.headers(headers)
        elif not SHEET_AVAILABLE:
            # fallback: show as CSV text
            self.text_sheet.delete(1.0, tk.END)
            self.text_sheet.insert(tk.END, self.df.to_csv(index=False, sep='\t'))

    # ============================================================
    # Graph Tab Construction
    # ============================================================
    def build_graph_tab(self):
        # Create matplotlib figure and canvas
        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Toolbar
        toolbar_frame = ttk.Frame(self.graph_frame)
        toolbar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()

        # Plot controls
        control_frame = ttk.Frame(self.graph_frame)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=2)
        ttk.Label(control_frame, text="X col:").pack(side=tk.LEFT, padx=2)
        self.x_col_var = tk.StringVar()
        self.x_col_combo = ttk.Combobox(control_frame, textvariable=self.x_col_var, width=10)
        self.x_col_combo.pack(side=tk.LEFT, padx=2)
        ttk.Label(control_frame, text="Y col(s):").pack(side=tk.LEFT, padx=2)
        self.y_cols_listbox = tk.Listbox(control_frame, selectmode=tk.MULTIPLE, height=3, width=10, exportselection=False)
        self.y_cols_listbox.pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="Plot", command=self.plot_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Clear", command=self.clear_graph).pack(side=tk.LEFT, padx=5)

        self.update_graph_column_choices()

    def update_graph_column_choices(self):
        cols = list(self.df.columns)
        self.x_col_combo['values'] = cols
        self.y_cols_listbox.delete(0, tk.END)
        for c in cols:
            self.y_cols_listbox.insert(tk.END, c)
        if cols:
            self.x_col_var.set(cols[0])

    # ============================================================
    # Data Loading & Sample
    # ============================================================
    def load_sample_data(self):
        t = np.linspace(0, 10, 101)
        v = 5 * np.sin(2 * np.pi * 0.5 * t) + 12 + np.random.normal(0, 0.5, len(t))
        i = 0.5 * np.sin(2 * np.pi * 0.5 * t + np.pi/4) + 2 + np.random.normal(0, 0.1, len(t))
        self.df = pd.DataFrame({'Time (s)': t, 'Voltage (V)': v, 'Current (A)': i})
        self.update_sheet_display()
        self.update_graph_column_choices()

    def new_workbook(self):
        if messagebox.askyesno("New", "Clear all data?"):
            self.df = pd.DataFrame(columns=['A','B','C'])
            self.update_sheet_display()
            self.update_graph_column_choices()
            self.clear_graph()
            self.undo_stack.clear()

    def open_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if path:
            try:
                self.df = pd.read_csv(path)
                self.update_sheet_display()
                self.update_graph_column_choices()
                self.undo_stack.clear()
                self.status.config(text=f"Loaded {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV:\n{e}")

    def save_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV Files", "*.csv")])
        if path:
            try:
                self.df.to_csv(path, index=False)
                self.status.config(text=f"Saved {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save:\n{e}")

    # ============================================================
    # Edit Operations (Undo/Redo, Row/Col insertion)
    # ============================================================
    def push_undo(self):
        self.undo_stack.append(self.df.copy())
        if len(self.undo_stack) > 50:  # limit
            self.undo_stack.pop(0)

    def undo(self):
        if self.undo_stack:
            self.df = self.undo_stack.pop()
            self.update_sheet_display()
            self.update_graph_column_choices()
        else:
            messagebox.showinfo("Undo", "Nothing to undo")

    def redo(self):
        # simplified: redo by keeping a redo stack (not implemented)
        messagebox.showinfo("Redo", "Redo not yet implemented")

    def insert_row(self):
        self.push_undo()
        self.df.loc[self.df.shape[0]] = [np.nan] * len(self.df.columns)
        self.update_sheet_display()

    def delete_row(self):
        self.push_undo()
        if self.sheet and SHEET_AVAILABLE:
            selected = self.sheet.get_selected_rows()
            if selected:
                self.df.drop(self.df.index[selected], inplace=True)
                self.df.reset_index(drop=True, inplace=True)
                self.update_sheet_display()
            else:
                messagebox.showwarning("Delete", "Select rows to delete.")
        else:
            messagebox.showwarning("Delete", "Row selection not available in fallback mode.")

    def insert_column(self):
        self.push_undo()
        col_name = simpledialog.askstring("Insert Column", "Column name:")
        if col_name:
            self.df[col_name] = np.nan
            self.update_sheet_display()
            self.update_graph_column_choices()

    def delete_column(self):
        self.push_undo()
        col_name = simpledialog.askstring("Delete Column", "Column name:")
        if col_name and col_name in self.df.columns:
            self.df.drop(columns=[col_name], inplace=True)
            self.update_sheet_display()
            self.update_graph_column_choices()

    def sort_column(self, ascending):
        self.push_undo()
        if self.sheet and SHEET_AVAILABLE:
            selected = self.sheet.get_selected_columns()
            if selected:
                col = self.df.columns[selected[0]]
                self.df.sort_values(by=col, ascending=ascending, inplace=True)
                self.df.reset_index(drop=True, inplace=True)
                self.update_sheet_display()
            else:
                messagebox.showwarning("Sort", "Select a column first.")
        else:
            messagebox.showwarning("Sort", "Not available in fallback mode.")

    def fill_down(self):
        self.push_undo()
        # Get selected range from tksheet
        if self.sheet and SHEET_AVAILABLE:
            try:
                # Get currently edited cell or selection
                rows = self.sheet.get_selected_rows()
                cols = self.sheet.get_selected_columns()
                if rows and cols:
                    top_val = self.df.iat[rows[0], cols[0]]
                    for r in rows[1:]:
                        self.df.iat[r, cols[0]] = top_val
                    self.update_sheet_display()
            except:
                messagebox.showwarning("Fill Down", "Select a range of cells first.")

    def fill_right(self):
        self.push_undo()
        if self.sheet and SHEET_AVAILABLE:
            try:
                rows = self.sheet.get_selected_rows()
                cols = self.sheet.get_selected_columns()
                if rows and cols:
                    left_val = self.df.iat[rows[0], cols[0]]
                    for c in cols[1:]:
                        self.df.iat[rows[0], c] = left_val
                    self.update_sheet_display()
            except:
                messagebox.showwarning("Fill Right", "Select a range of cells first.")

    # ============================================================
    # Keyboard Shortcuts
    # ============================================================
    def bind_shortcuts(self):
        self.bind_all("<Control-n>", lambda e: self.new_workbook())
        self.bind_all("<Control-o>", lambda e: self.open_csv())
        self.bind_all("<Control-s>", lambda e: self.save_csv())
        self.bind_all("<Control-z>", lambda e: self.undo())
        self.bind_all("<Control-p>", lambda e: self.plot_selected())

    # ============================================================
    # Graph Operations
    # ============================================================
    def plot_selected(self):
        x_col = self.x_col_var.get()
        y_indices = self.y_cols_listbox.curselection()
        if not x_col or not y_indices:
            messagebox.showwarning("Plot", "Select X and at least one Y column.")
            return
        y_cols = [self.y_cols_listbox.get(i) for i in y_indices]
        self.ax.clear()
        for y_col in y_cols:
            if x_col in self.df.columns and y_col in self.df.columns:
                x = pd.to_numeric(self.df[x_col], errors='coerce')
                y = pd.to_numeric(self.df[y_col], errors='coerce')
                mask = x.notna() & y.notna()
                self.ax.plot(x[mask], y[mask], marker='o', markersize=3, label=y_col)
        self.ax.set_xlabel(x_col)
        self.ax.set_ylabel(", ".join(y_cols))
        self.ax.legend()
        self.ax.grid(True, linestyle='--', alpha=0.7)
        self.fig.tight_layout()
        self.canvas.draw()

    def clear_graph(self):
        self.ax.clear()
        self.canvas.draw()

    # ============================================================
    # Analysis Functions
    # ============================================================
    def show_descriptive_stats(self):
        stats_df = self.df.describe(include='all').T
        stats_df.insert(0, 'Column', stats_df.index)
        # Show in a new window
        win = tk.Toplevel(self)
        win.title("Descriptive Statistics")
        text = ScrolledText(win, wrap=tk.NONE, width=100, height=20)
        text.pack(fill=tk.BOTH, expand=True)
        text.insert(tk.END, stats_df.to_string(index=False))

    def curve_fitting_dialog(self):
        # Simple dialog for linear/poly fit
        cols = list(self.df.columns)
        if len(cols) < 2:
            messagebox.showwarning("Fit", "Need at least two columns.")
            return
        win = tk.Toplevel(self)
        win.title("Curve Fitting")
        ttk.Label(win, text="X Column:").grid(row=0, column=0, padx=5, pady=5)
        x_var = tk.StringVar(value=cols[0])
        x_combo = ttk.Combobox(win, textvariable=x_var, values=cols, state='readonly')
        x_combo.grid(row=0, column=1)
        ttk.Label(win, text="Y Column:").grid(row=1, column=0, padx=5, pady=5)
        y_var = tk.StringVar(value=cols[1])
        y_combo = ttk.Combobox(win, textvariable=y_var, values=cols, state='readonly')
        y_combo.grid(row=1, column=1)
        ttk.Label(win, text="Degree:").grid(row=2, column=0, padx=5, pady=5)
        deg_var = tk.IntVar(value=1)
        deg_spin = ttk.Spinbox(win, from_=1, to=10, textvariable=deg_var, width=5)
        deg_spin.grid(row=2, column=1)

        def do_fit():
            x_col = x_var.get()
            y_col = y_var.get()
            x = pd.to_numeric(self.df[x_col], errors='coerce')
            y = pd.to_numeric(self.df[y_col], errors='coerce')
            mask = x.notna() & y.notna()
            x = x[mask].values
            y = y[mask].values
            coeffs = np.polyfit(x, y, deg_var.get())
            poly = np.poly1d(coeffs)
            # Plot
            self.ax.clear()
            self.ax.scatter(x, y, s=10, label='Data')
            x_line = np.linspace(min(x), max(x), 200)
            self.ax.plot(x_line, poly(x_line), 'r-', label=f'Poly fit (deg={deg_var.get()})')
            self.ax.legend()
            self.ax.set_xlabel(x_col)
            self.ax.set_ylabel(y_col)
            self.fig.tight_layout()
            self.canvas.draw()
            # Show equation
            eq_str = " + ".join(f"{c:.4g}*x^{i}" for i, c in enumerate(reversed(coeffs)))
            messagebox.showinfo("Fit Result", f"Equation:\n{eq_str}")
            win.destroy()

        ttk.Button(win, text="Fit & Plot", command=do_fit).grid(row=3, column=0, columnspan=2, pady=10)

    def show_fft(self):
        cols = list(self.df.columns)
        if len(cols) < 2:
            messagebox.showwarning("FFT", "Need at least two columns.")
            return
        win = tk.Toplevel(self)
        win.title("FFT")
        ttk.Label(win, text="X Column:").grid(row=0, column=0, padx=5, pady=5)
        x_var = tk.StringVar(value=cols[0])
        x_combo = ttk.Combobox(win, textvariable=x_var, values=cols, state='readonly')
        x_combo.grid(row=0, column=1)
        ttk.Label(win, text="Y Column:").grid(row=1, column=0, padx=5, pady=5)
        y_var = tk.StringVar(value=cols[1])
        y_combo = ttk.Combobox(win, textvariable=y_var, values=cols, state='readonly')
        y_combo.grid(row=1, column=1)

        def do_fft():
            x_col = x_var.get()
            y_col = y_var.get()
            x = pd.to_numeric(self.df[x_col], errors='coerce')
            y = pd.to_numeric(self.df[y_col], errors='coerce')
            mask = x.notna() & y.notna()
            x = x[mask].values
            y = y[mask].values
            N = len(y)
            T = x[1] - x[0] if len(x) > 1 else 1.0
            yf = fft(y)
            xf = fftfreq(N, T)[:N//2]
            amplitude = 2.0/N * np.abs(yf[:N//2])
            self.ax.clear()
            self.ax.plot(xf, amplitude)
            self.ax.set_xlabel('Frequency (Hz)')
            self.ax.set_ylabel('Amplitude')
            self.ax.set_title('FFT')
            self.fig.tight_layout()
            self.canvas.draw()
            win.destroy()

        ttk.Button(win, text="Compute FFT", command=do_fft).grid(row=2, column=0, columnspan=2, pady=10)

    def show_histogram(self):
        cols = list(self.df.columns)
        if not cols:
            return
        win = tk.Toplevel(self)
        win.title("Histogram")
        ttk.Label(win, text="Column:").grid(row=0, column=0, padx=5, pady=5)
        col_var = tk.StringVar(value=cols[0])
        col_combo = ttk.Combobox(win, textvariable=col_var, values=cols, state='readonly')
        col_combo.grid(row=0, column=1)
        ttk.Label(win, text="Bins:").grid(row=1, column=0, padx=5, pady=5)
        bins_var = tk.IntVar(value=10)
        bins_spin = ttk.Spinbox(win, from_=5, to=100, textvariable=bins_var, width=5)
        bins_spin.grid(row=1, column=1)

        def do_hist():
            col = col_var.get()
            data = pd.to_numeric(self.df[col], errors='coerce').dropna().values
            self.ax.clear()
            self.ax.hist(data, bins=bins_var.get(), edgecolor='black', alpha=0.7)
            self.ax.set_xlabel(col)
            self.ax.set_ylabel('Frequency')
            self.ax.set_title('Histogram')
            self.fig.tight_layout()
            self.canvas.draw()
            win.destroy()

        ttk.Button(win, text="Plot Histogram", command=do_hist).grid(row=2, column=0, columnspan=2, pady=10)


# ============================================================
# Run Application
# ============================================================
if __name__ == "__main__":
    app = DataGraphApp()
    app.mainloop()