import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from db_loader import load_excel_to_pickle
from generate_reports import generate_pivot_report
import threading

class DataApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Loader and Report Generator")
        self.geometry("900x700")
        self.style = ttk.Style(self)
        self.style.theme_use('clam')

        # Theme colors
        self.light_theme = {
            "background": "#f0f0f0",
            "foreground": "#000000",
            "button_background": "#e1e1e1",
            "button_foreground": "#000000",
            "entry_background": "#ffffff",
            "text_background": "#ffffff",
            "text_foreground": "#000000",
            "listbox_background": "#ffffff",
            "listbox_foreground": "#000000",
            "highlight_color": "#0078d7"
        }
        self.dark_theme = {
            "background": "#2e2e2e",
            "foreground": "#ffffff",
            "button_background": "#444444",
            "button_foreground": "#ffffff",
            "entry_background": "#3c3c3c",
            "text_background": "#3c3c3c",
            "text_foreground": "#ffffff",
            "listbox_background": "#3c3c3c",
            "listbox_foreground": "#ffffff",
            "highlight_color": "#3399ff"
        }
        self.current_theme = "light"

        self.data_dir = "./data"
        self.df = None
        self.pkl_files = []
        self.filter_columns = []
        self.filter_values = {}
        self.selected_display_columns = []

        self.create_widgets()
        self.create_theme_toggle()
        self.apply_theme()

    def create_widgets(self):
        # Create main menu frame
        self.main_menu_frame = ttk.Frame(self, padding=20)
        self.main_menu_frame.pack(fill='both', expand=True)

        ttk.Label(self.main_menu_frame, text="Main Menu", font=("Arial", 18)).pack(pady=20)

        ttk.Button(self.main_menu_frame, text="Load Data", command=self.show_load_tab).pack(pady=10, ipadx=10, ipady=5)
        ttk.Button(self.main_menu_frame, text="Select Data", command=self.show_select_tab).pack(pady=10, ipadx=10, ipady=5)
        ttk.Button(self.main_menu_frame, text="Generate Report", command=self.show_report_tab).pack(pady=10, ipadx=10, ipady=5)

        # Create frames for each section but do not pack yet
        self.tab_load = ttk.Frame(self, padding=20)
        self.tab_select = ttk.Frame(self, padding=20)
        self.tab_report = ttk.Frame(self, padding=10)

        self.create_load_tab()
        self.create_select_tab()
        self.create_report_tab()

        # Initially hide all section frames
        self.tab_load.pack_forget()
        self.tab_select.pack_forget()
        self.tab_report.pack_forget()

    def show_main_menu(self):
        self.main_menu_frame.pack(fill='both', expand=True)
        self.tab_load.pack_forget()
        self.tab_select.pack_forget()
        self.tab_report.pack_forget()

    def show_load_tab(self):
        self.main_menu_frame.pack_forget()
        self.tab_load.pack(fill='both', expand=True)
        self.add_back_button(self.tab_load)

    def show_select_tab(self):
        self.main_menu_frame.pack_forget()
        self.tab_select.pack(fill='both', expand=True)
        self.add_back_button(self.tab_select)

    def show_report_tab(self):
        self.main_menu_frame.pack_forget()
        self.tab_report.pack(fill='both', expand=True)
        self.add_back_button(self.tab_report)
        self.prepare_report_tab()

    def add_back_button(self, parent_frame):
        # Remove existing back button if any
        for child in parent_frame.winfo_children():
            if getattr(child, 'is_back_button', False):
                child.destroy()
        back_btn = ttk.Button(parent_frame, text="Back to Main Menu", command=self.show_main_menu)
        back_btn.is_back_button = True
        back_btn.pack(anchor='ne', pady=5, padx=5)

    def update_filter_value_entries(self):
        # Clear previous entries
        for widget in self.filter_values_container.winfo_children():
            widget.destroy()
        self.filter_entries.clear()

        selected_indices = self.filter_listbox.curselection()
        for idx in selected_indices:
            col = self.filter_listbox.get(idx)
            label = ttk.Label(self.filter_values_container, text=f"Values for {col}:")
            label.pack(anchor='w', padx=5, pady=2)

            # Create listbox with unique values from dataframe column for multiple selection
            values = []
            if self.df is not None and col in self.df.columns:
                values = sorted(self.df[col].dropna().astype(str).unique().tolist())
            listbox = tk.Listbox(self.filter_values_container, selectmode='multiple', exportselection=0, height=min(10, len(values)))
            for val in values:
                listbox.insert(tk.END, val)
            listbox.pack(fill='x', padx=5, pady=2)
            self.filter_entries[col] = listbox

        selected_indices = self.filter_listbox.curselection()
        for idx in selected_indices:
            col = self.filter_listbox.get(idx)
            label = ttk.Label(self.filter_values_container, text=f"Value for {col}:")
            label.pack(anchor='w', padx=5, pady=2)

            # Create combobox with unique values from dataframe column
            values = []
            if self.df is not None and col in self.df.columns:
                values = sorted(self.df[col].dropna().astype(str).unique().tolist())
            combobox = ttk.Combobox(self.filter_values_container, values=values, state='readonly')
            combobox.pack(fill='x', padx=5, pady=2)
            if values:
                combobox.current(0)
            self.filter_entries[col] = combobox

    def generate_text_report(self):
        if self.df is None:
            messagebox.showerror("Error", "No DataFrame loaded.")
            return
        filter_cols = [self.filter_listbox.get(i) for i in self.filter_listbox.curselection()]
        if not filter_cols:
            messagebox.showerror("Error", "Select at least one filter column.")
            return
        criteria = {}
        for col in filter_cols:
            val = self.filter_entries.get(col)
            if val:
                v = val.get().strip()
                if v == "":
                    messagebox.showerror("Error", f"Filter value for '{col}' is empty.")
                    return
                criteria[col] = v
            else:
                messagebox.showerror("Error", f"Filter value for '{col}' is missing.")
                return
        display_cols = [self.display_listbox.get(i) for i in self.display_listbox.curselection()]
        if not display_cols:
            messagebox.showerror("Error", "Select at least one display column.")
            return

        mask = pd.Series(True, index=self.df.index)
        for col, val in criteria.items():
            mask &= self.df[col].astype(str) == val

        result = self.df.loc[mask, display_cols]
        self.report_text.config(state='normal')
        self.report_text.delete('1.0', tk.END)
        if result.empty:
            self.report_text.insert(tk.END, "No data matching the filters.\n")
        else:
            self.report_text.insert(tk.END, result.to_string())
        self.report_text.config(state='disabled')

    def create_theme_toggle(self):
        # Add a theme toggle button at the top right corner
        self.theme_var = tk.StringVar(value=self.current_theme)
        toggle_frame = ttk.Frame(self)
        toggle_frame.place(relx=1.0, rely=0.0, anchor='ne', x=-10, y=10)
        self.theme_button = ttk.Button(toggle_frame, text="Switch to Dark Theme", command=self.toggle_theme)
        self.theme_button.pack()

    def toggle_theme(self):
        if self.current_theme == "light":
            self.current_theme = "dark"
            self.theme_button.config(text="Switch to Light Theme")
        else:
            self.current_theme = "light"
            self.theme_button.config(text="Switch to Dark Theme")
        self.apply_theme()

    def apply_theme(self):
        theme = self.light_theme if self.current_theme == "light" else self.dark_theme

        # Configure overall background
        self.configure(bg=theme["background"])

        # Style configuration for ttk widgets
        self.style.configure('TFrame', background=theme["background"])
        self.style.configure('TLabel', background=theme["background"], foreground=theme["foreground"])
        self.style.configure('TButton', background=theme["button_background"], foreground=theme["button_foreground"])
        self.style.map('TButton',
                       background=[('active', theme["highlight_color"])],
                       foreground=[('active', theme["button_foreground"])])
        self.style.configure('TEntry', fieldbackground=theme["entry_background"], foreground=theme["foreground"])
        self.style.configure('TCombobox', fieldbackground=theme["entry_background"], foreground=theme["foreground"])
        self.style.configure('TNotebook', background=theme["background"])
        self.style.configure('TNotebook.Tab', background=theme["button_background"], foreground=theme["foreground"])
        self.style.map('TNotebook.Tab',
                       background=[('selected', theme["highlight_color"])],
                       foreground=[('selected', theme["foreground"])])

        # Update widgets background and foreground colors
        def recursive_configure(widget):
            for child in widget.winfo_children():
                cls = child.winfo_class()
                # Skip ttk widgets for direct config, style them via ttk.Style
                if cls in ['Frame', 'Label', 'Button', 'Entry', 'Text', 'Listbox']:
                    if cls == 'Frame':
                        child.configure(background=theme["background"])
                    elif cls == 'Label':
                        child.configure(background=theme["background"], foreground=theme["foreground"])
                    elif cls == 'Button':
                        # ttk buttons styled by style, skip
                        pass
                    elif cls == 'Entry':
                        try:
                            child.configure(background=theme["entry_background"], foreground=theme["foreground"])
                        except:
                            pass
                    elif cls == 'Text':
                        child.configure(background=theme["text_background"], foreground=theme["text_foreground"], insertbackground=theme["foreground"])
                    elif cls == 'Listbox':
                        child.configure(background=theme["listbox_background"], foreground=theme["listbox_foreground"], selectbackground=theme["highlight_color"])
                recursive_configure(child)

        recursive_configure(self)

    def create_load_tab(self):
        frame = ttk.Frame(self.tab_load, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Load Excel Sheets to Pickle Files", font=("Arial", 14)).pack(pady=10)

        self.excel_path_var = tk.StringVar(value="DZ_2.xlsx")
        entry = ttk.Entry(frame, textvariable=self.excel_path_var, width=50)
        entry.pack(pady=5)

        def browse_file():
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            if file_path:
                self.excel_path_var.set(file_path)

        ttk.Button(frame, text="Browse Excel File", command=browse_file).pack(pady=5)
        ttk.Button(frame, text="Load Excel to Pickle", command=self.load_excel_thread).pack(pady=10)

        self.load_status = ttk.Label(frame, text="")
        self.load_status.pack(pady=5)

    def load_excel_thread(self):
        threading.Thread(target=self.load_excel_action, daemon=True).start()

    def load_excel_action(self):
        excel_path = self.excel_path_var.get()
        if not os.path.isfile(excel_path):
            self.show_status("Excel file not found.", error=True)
            return
        try:
            load_excel_to_pickle(excel_path, self.data_dir)
            self.show_status("Excel sheets loaded to pickle files successfully.")
            self.refresh_pkl_files()
        except Exception as e:
            self.show_status(f"Error loading Excel: {e}", error=True)

    def show_status(self, message, error=False):
        self.load_status.config(text=message, foreground='red' if error else 'green')

    def create_select_tab(self):
        frame = ttk.Frame(self.tab_select, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Select Pickle File to Load DataFrame", font=("Arial", 14)).pack(pady=10)

        self.pkl_var = tk.StringVar()
        self.pkl_combo = ttk.Combobox(frame, textvariable=self.pkl_var, state='readonly', width=50)
        self.pkl_combo.pack(pady=5)
        self.refresh_pkl_files()

        ttk.Button(frame, text="Load DataFrame", command=self.load_dataframe).pack(pady=10)

        self.df_info = tk.Text(frame, height=15, width=80, state='disabled', wrap='none')
        self.df_info.pack(pady=10, fill='both', expand=True)

        # Add scrollbars for df_info
        yscroll = ttk.Scrollbar(frame, orient='vertical', command=self.df_info.yview)
        yscroll.pack(side='right', fill='y')
        self.df_info['yscrollcommand'] = yscroll.set

        xscroll = ttk.Scrollbar(frame, orient='horizontal', command=self.df_info.xview)
        xscroll.pack(side='bottom', fill='x')
        self.df_info['xscrollcommand'] = xscroll.set

    def refresh_pkl_files(self):
        if os.path.isdir(self.data_dir):
            self.pkl_files = [f for f in os.listdir(self.data_dir) if f.endswith('.pkl')]
        else:
            self.pkl_files = []
        self.pkl_combo['values'] = self.pkl_files
        if self.pkl_files:
            self.pkl_combo.current(0)

    def load_dataframe(self):
        selected_file = self.pkl_var.get()
        if not selected_file:
            messagebox.showerror("Error", "Please select a pickle file.")
            return
        try:
            path = os.path.join(self.data_dir, selected_file)
            self.df = pd.read_pickle(path)
            self.df.columns = self.df.columns.str.strip()
            self.show_dataframe_info()
            self.prepare_report_tab()
            messagebox.showinfo("Success", f"DataFrame loaded from {selected_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load DataFrame: {e}")

    def show_dataframe_info(self):
        self.df_info.config(state='normal')
        self.df_info.delete('1.0', tk.END)
        info_text = f"Columns:\n"
        for i, col in enumerate(self.df.columns, 1):
            info_text += f"{i}. {col}\n"
        info_text += "\nFirst 5 rows:\n"
        info_text += self.df.head().to_string()
        self.df_info.insert(tk.END, info_text)
        self.df_info.config(state='disabled')

    def create_report_tab(self):
        frame = ttk.Frame(self.tab_report, padding=10)
        frame.pack(fill='both', expand=True)

        # Filtering columns selection
        filter_frame = ttk.LabelFrame(frame, text="Filter Columns")
        filter_frame.grid(row=0, column=0, sticky='nswe', padx=5, pady=5)

        self.filter_listbox = tk.Listbox(filter_frame, selectmode='multiple', exportselection=0, height=10)
        self.filter_listbox.pack(side='left', fill='both', expand=True, padx=(5,0), pady=5)
        self.filter_listbox.bind('<<ListboxSelect>>', lambda e: self.update_filter_value_entries())

        filter_scroll = ttk.Scrollbar(filter_frame, orient='vertical', command=self.filter_listbox.yview)
        filter_scroll.pack(side='right', fill='y')
        self.filter_listbox.config(yscrollcommand=filter_scroll.set)

        # Filter values input
        value_frame = ttk.LabelFrame(frame, text="Filter Values")
        value_frame.grid(row=0, column=1, sticky='nswe', padx=5, pady=5)

        self.filter_entries = {}

        self.filter_values_container = ttk.Frame(value_frame)
        self.filter_values_container.pack(fill='both', expand=True, padx=5, pady=5)

        # Display columns selection
        display_frame = ttk.LabelFrame(frame, text="Display Columns")
        display_frame.grid(row=1, column=0, sticky='nswe', padx=5, pady=5)

        self.display_listbox = tk.Listbox(display_frame, selectmode='multiple', exportselection=0, height=10)
        self.display_listbox.pack(side='left', fill='both', expand=True, padx=(5,0), pady=5)

        display_scroll = ttk.Scrollbar(display_frame, orient='vertical', command=self.display_listbox.yview)
        display_scroll.pack(side='right', fill='y')
        self.display_listbox.config(yscrollcommand=display_scroll.set)

        # Report buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=1, column=1, sticky='nswe', padx=5, pady=5)

        ttk.Button(button_frame, text="Generate Text Report", command=self.generate_text_report).pack(fill='x', pady=3)
        ttk.Button(button_frame, text="Generate Scatter Plot", command=self.generate_scatter_plot).pack(fill='x', pady=3)
        ttk.Button(button_frame, text="Generate Pie Chart", command=self.generate_pie_chart).pack(fill='x', pady=3)
        ttk.Button(button_frame, text="Generate Bar Chart", command=self.generate_bar_chart).pack(fill='x', pady=3)
        ttk.Button(button_frame, text="Generate Pivot Table", command=self.generate_pivot_report).pack(fill='x', pady=3)

        # Text report output
        output_frame = ttk.LabelFrame(frame, text="Report Output")
        output_frame.grid(row=2, column=0, columnspan=2, sticky='nswe', padx=5, pady=5)

        self.report_text = tk.Text(output_frame, height=15, wrap='none')
        self.report_text.pack(fill='both', expand=True, padx=5, pady=5)

        # Scrollbars for report_text
        yscroll = ttk.Scrollbar(output_frame, orient='vertical', command=self.report_text.yview)
        yscroll.pack(side='right', fill='y')
        self.report_text['yscrollcommand'] = yscroll.set

        xscroll = ttk.Scrollbar(output_frame, orient='horizontal', command=self.report_text.xview)
        xscroll.pack(side='bottom', fill='x')
        self.report_text['xscrollcommand'] = xscroll.set

        # Configure grid weights
        frame.rowconfigure(2, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

    def prepare_report_tab(self):
        if self.df is None:
            return
        cols = self.df.columns.tolist()
        self.filter_listbox.delete(0, tk.END)
        self.display_listbox.delete(0, tk.END)
        for col in cols:
            self.filter_listbox.insert(tk.END, col)
            self.display_listbox.insert(tk.END, col)
        self.update_filter_value_entries()

    def update_filter_value_entries(self):
        # Clear previous entries
        for widget in self.filter_values_container.winfo_children():
            widget.destroy()
        self.filter_entries.clear()

        selected_indices = self.filter_listbox.curselection()
        for idx in selected_indices:
            col = self.filter_listbox.get(idx)
            label = ttk.Label(self.filter_values_container, text=f"Value for {col}:")
            label.pack(anchor='w', padx=5, pady=2)

            # Create combobox with unique values from dataframe column
            values = []
            if self.df is not None and col in self.df.columns:
                values = sorted(self.df[col].dropna().astype(str).unique().tolist())
            combobox = ttk.Combobox(self.filter_values_container, values=values, state='readonly')
            combobox.pack(fill='x', padx=5, pady=2)
            if values:
                combobox.current(0)
            self.filter_entries[col] = combobox

    def generate_text_report(self):
        if self.df is None:
            messagebox.showerror("Error", "No DataFrame loaded.")
            return
        filter_cols = [self.filter_listbox.get(i) for i in self.filter_listbox.curselection()]
        if not filter_cols:
            messagebox.showerror("Error", "Select at least one filter column.")
            return
        criteria = {}
        for col in filter_cols:
            val = self.filter_entries.get(col)
            if val:
                v = val.get().strip()
                if v == "":
                    messagebox.showerror("Error", f"Filter value for '{col}' is empty.")
                    return
                criteria[col] = v
            else:
                messagebox.showerror("Error", f"Filter value for '{col}' is missing.")
                return
        display_cols = [self.display_listbox.get(i) for i in self.display_listbox.curselection()]
        if not display_cols:
            messagebox.showerror("Error", "Select at least one display column.")
            return

        mask = pd.Series(True, index=self.df.index)
        for col, val in criteria.items():
            mask &= self.df[col].astype(str) == val

        result = self.df.loc[mask, display_cols]
        self.report_text.config(state='normal')
        self.report_text.delete('1.0', tk.END)
        if result.empty:
            self.report_text.insert(tk.END, "No data matching the filters.\n")
        else:
            self.report_text.insert(tk.END, result.to_string())
        self.report_text.config(state='disabled')

    def generate_scatter_plot(self):
        if self.df is None:
            messagebox.showerror("Error", "No DataFrame loaded.")
            return
        cols = self.df.columns.tolist()
        ScatterDialog(self, self.df, cols)

    def generate_pie_chart(self):
        if self.df is None:
            messagebox.showerror("Error", "No DataFrame loaded.")
            return
        cols = self.df.columns.tolist()
        PieDialog(self, self.df, cols)

    def generate_bar_chart(self):
        if self.df is None:
            messagebox.showerror("Error", "No DataFrame loaded.")
            return
        cols = self.df.columns.tolist()
        BarDialog(self, self.df, cols)

    def generate_pivot_report(self):
        if self.df is None:
            messagebox.showerror("Error", "No DataFrame loaded.")
            return
        PivotDialog(self, self.df)

class ScatterDialog(tk.Toplevel):
    def __init__(self, parent, df, columns):
        super().__init__(parent)
        self.title("Scatter Plot")
        self.df = df
        self.columns = columns
        self.geometry("600x500")

        ttk.Label(self, text="Select X column:").pack(pady=5)
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(self, textvariable=self.x_var, values=columns, state='readonly')
        self.x_combo.pack(pady=5)
        self.x_combo.current(0)

        ttk.Label(self, text="Select Y column:").pack(pady=5)
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(self, textvariable=self.y_var, values=columns, state='readonly')
        self.y_combo.pack(pady=5)
        self.y_combo.current(0)

        ttk.Button(self, text="Plot", command=self.plot).pack(pady=10)

        self.fig, self.ax = plt.subplots(figsize=(6,4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot(self):
        x = self.x_var.get()
        y = self.y_var.get()
        self.ax.clear()
        try:
            self.df.plot.scatter(x=x, y=y, ax=self.ax)
            self.ax.set_title(f"{y} vs {x}")
            self.ax.grid(True)
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot scatter: {e}")

class PieDialog(tk.Toplevel):
    def __init__(self, parent, df, columns):
        super().__init__(parent)
        self.title("Pie Chart")
        self.df = df
        self.columns = columns
        self.geometry("600x500")

        ttk.Label(self, text="Select column for Pie Chart:").pack(pady=5)
        self.col_var = tk.StringVar()
        self.col_combo = ttk.Combobox(self, textvariable=self.col_var, values=columns, state='readonly')
        self.col_combo.pack(pady=5)
        self.col_combo.current(0)

        ttk.Button(self, text="Plot", command=self.plot).pack(pady=10)

        self.fig, self.ax = plt.subplots(figsize=(6,6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot(self):
        col = self.col_var.get()
        self.ax.clear()
        try:
            value_counts = self.df[col].value_counts()
            if value_counts.empty:
                messagebox.showinfo("Info", "No data for pie chart.")
                return
            value_counts.plot.pie(autopct='%1.1f%%', startangle=360, shadow=True, ax=self.ax)
            self.ax.set_title(f'Distribution by {col}')
            self.ax.set_ylabel('')
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot pie chart: {e}")

class BarDialog(tk.Toplevel):
    def __init__(self, parent, df, columns):
        super().__init__(parent)
        self.title("Bar Chart")
        self.df = df
        self.columns = columns
        self.geometry("600x500")

        ttk.Label(self, text="Select column for Bar Chart:").pack(pady=5)
        self.col_var = tk.StringVar()
        self.col_combo = ttk.Combobox(self, textvariable=self.col_var, values=columns, state='readonly')
        self.col_combo.pack(pady=5)
        self.col_combo.current(0)

        ttk.Button(self, text="Plot", command=self.plot).pack(pady=10)

        self.fig, self.ax = plt.subplots(figsize=(6,4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

    def plot(self):
        col = self.col_var.get()
        self.ax.clear()
        try:
            value_counts = self.df[col].value_counts()
            if value_counts.empty:
                messagebox.showinfo("Info", "No data for bar chart.")
                return
            value_counts.plot.bar(ax=self.ax)
            self.ax.set_title(f'Distribution by {col}')
            self.ax.set_ylabel('Count')
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot bar chart: {e}")

class PivotDialog(tk.Toplevel):
    def __init__(self, parent, df):
        super().__init__(parent)
        self.title("Pivot Table")
        self.df = df
        self.geometry("700x600")

        cols = df.columns.tolist()

        ttk.Label(self, text="Select index column:").pack(pady=5)
        self.index_var = tk.StringVar()
        self.index_combo = ttk.Combobox(self, textvariable=self.index_var, values=cols, state='readonly')
        self.index_combo.pack(pady=5)
        self.index_combo.current(0)

        ttk.Label(self, text="Select columns column:").pack(pady=5)
        self.columns_var = tk.StringVar()
        self.columns_combo = ttk.Combobox(self, textvariable=self.columns_var, values=cols, state='readonly')
        self.columns_combo.pack(pady=5)
        self.columns_combo.current(0)

        ttk.Label(self, text="Select values column (optional):").pack(pady=5)
        self.values_var = tk.StringVar()
        self.values_combo = ttk.Combobox(self, textvariable=self.values_var, values=[""] + cols, state='readonly')
        self.values_combo.pack(pady=5)
        self.values_combo.current(0)

        ttk.Label(self, text="Aggregation function (sum, count, mean, size, etc.):").pack(pady=5)
        self.agg_entry = ttk.Entry(self)
        self.agg_entry.pack(pady=5)
        self.agg_entry.insert(0, "sum")

        ttk.Button(self, text="Generate Pivot Table", command=self.generate_pivot).pack(pady=10)

        self.output_text = tk.Text(self, height=20, wrap='none')
        self.output_text.pack(fill='both', expand=True, padx=5, pady=5)

        yscroll = ttk.Scrollbar(self, orient='vertical', command=self.output_text.yview)
        yscroll.pack(side='right', fill='y')
        self.output_text['yscrollcommand'] = yscroll.set

        xscroll = ttk.Scrollbar(self, orient='horizontal', command=self.output_text.xview)
        xscroll.pack(side='bottom', fill='x')
        self.output_text['xscrollcommand'] = xscroll.set

    def generate_pivot(self):
        index_col = self.index_var.get()
        columns_col = self.columns_var.get()
        values_col = self.values_var.get() or None
        aggfunc = self.agg_entry.get().strip()

        try:
            pivot = pd.pivot_table(
                self.df,
                index=index_col,
                columns=columns_col,
                values=values_col,
                aggfunc=aggfunc,
                fill_value=0,
            )
            self.output_text.config(state='normal')
            self.output_text.delete('1.0', tk.END)
            self.output_text.insert(tk.END, str(pivot))
            self.output_text.config(state='disabled')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate pivot table: {e}")

if __name__ == "__main__":
    app = DataApp()
    app.mainloop()
