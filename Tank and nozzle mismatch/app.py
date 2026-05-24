import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import datetime

# --- Logic Functions ---
def categorize_pump_status(value):
    if pd.isna(value) or value == '' or str(value).lower() == 'nil':
        return 'Inactive'
    try:
        numeric_value = float(value)
        return 'Active' if numeric_value > 0 else 'Inactive'
    except (ValueError, TypeError):
        return 'Inactive'

class IndianOilAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Indian Oil RO Online Communicator")
        self.root.geometry("1400x950")
        self.root.configure(bg="#001a33")

        self.status_colors = {'Active': '#28a745', 'Inactive': '#dc3545'}
        self.date_column = "Last Data Recvd on (Date)"
        self.original_df = None
        self.current_df = None # Track what is currently in the table
        self.total_active_master = 0

        # --- Main Scrollable Container ---
        self.canvas = tk.Canvas(root, bg="#001a33", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#001a33")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.root.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self.setup_ui()

    def setup_ui(self):
        # 1. Header
        tk.Label(self.scrollable_frame, text="💾 Indian Oil RO Online Communicator", 
                 font=("Arial", 28, "bold"), fg="#ff851b", bg="#001a33").pack(pady=20)

        # 2. Top Buttons
        btn_frame = tk.Frame(self.scrollable_frame, bg="#001a33")
        btn_frame.pack(pady=10)

        self.upload_btn = tk.Button(btn_frame, text="1. Upload Sales Data", command=self.upload_file, 
                                   bg="#ff851b", fg="white", font=("Arial", 11, "bold"), width=18)
        self.upload_btn.pack(side=tk.LEFT, padx=5)

        self.match_btn = tk.Button(btn_frame, text="2. Load 2nd XL & Calculate", command=self.upload_mapping, 
                                  bg="#007bff", fg="white", font=("Arial", 11, "bold"), width=22, state="disabled")
        self.match_btn.pack(side=tk.LEFT, padx=5)

        self.download_btn = tk.Button(btn_frame, text="3. Download Table (.xlsx)", command=self.download_table, 
                                     bg="#28a745", fg="white", font=("Arial", 11, "bold"), width=22, state="disabled")
        self.download_btn.pack(side=tk.LEFT, padx=5)

        self.all_btn = tk.Button(btn_frame, text="Reset View", command=self.show_all, 
                                bg="#17a2b8", fg="white", font=("Arial", 11, "bold"), width=12, state="disabled")
        self.all_btn.pack(side=tk.LEFT, padx=5)

        # 3. Filter Section
        self.filter_frame = tk.LabelFrame(self.scrollable_frame, text=" Date Filter ", 
                                          font=("Arial", 10, "bold"), fg="#ff851b", bg="#001a33", padx=10, pady=15)
        self.filter_frame.pack(pady=10, padx=40, fill=tk.X)

        tk.Label(self.filter_frame, text="Day:", fg="white", bg="#001a33").grid(row=0, column=0, padx=5)
        self.day_cb = ttk.Combobox(self.filter_frame, width=8, state="disabled")
        self.day_cb.grid(row=0, column=1, padx=5)

        tk.Label(self.filter_frame, text="Month:", fg="white", bg="#001a33").grid(row=0, column=2, padx=5)
        self.month_cb = ttk.Combobox(self.filter_frame, width=12, state="disabled")
        self.month_cb.grid(row=0, column=3, padx=5)

        tk.Label(self.filter_frame, text="Year:", fg="white", bg="#001a33").grid(row=0, column=4, padx=5)
        self.year_cb = ttk.Combobox(self.filter_frame, width=8, state="disabled")
        self.year_cb.grid(row=0, column=5, padx=5)

        self.filter_btn = tk.Button(self.filter_frame, text="Apply Filter", command=self.apply_filter, 
                                   bg="#28a745", fg="white", font=("Arial", 10, "bold"), state="disabled")
        self.filter_btn.grid(row=0, column=6, padx=20)

        # 4. Summary & Charts
        self.summary_label = tk.Label(self.scrollable_frame, text="", font=("Arial", 18, "bold"), fg="white", bg="#001a33")
        self.summary_label.pack(pady=15)

        self.charts_frame = tk.Frame(self.scrollable_frame, bg="#001a33")
        self.charts_frame.pack(fill=tk.X, padx=20)

        # 5. Calculation Section
        self.calc_container = tk.LabelFrame(self.scrollable_frame, text=" Division Market Share Calculation ", 
                                            fg="#17a2b8", bg="#001a33", font=("Arial", 12, "bold"))
        self.calc_container.pack(fill=tk.X, padx=40, pady=20)
        
        self.calc_info_label = tk.Label(self.calc_container, text="Awaiting data...", 
                                       font=("Arial", 11, "italic"), fg="white", bg="#001a33", pady=10)
        self.calc_info_label.pack()

        # 6. Table Section
        self.table_main_container = tk.Frame(self.scrollable_frame, bg="#001a33", bd=2, relief=tk.SUNKEN)
        self.table_main_container.pack(fill=tk.BOTH, expand=True, padx=40, pady=(10, 50))
        
        self.table_canvas = tk.Frame(self.table_main_container, bg="#001a33")
        self.table_canvas.pack(fill=tk.BOTH, expand=True)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
        if not path: return
        try:
            df = pd.read_excel(path)
            df['Status'] = df['SAP Last 3 Months Sale'].apply(categorize_pump_status)
            df[self.date_column] = pd.to_datetime(df[self.date_column], errors='coerce')
            self.original_df = df.dropna(subset=[self.date_column])
            self.total_active_master = len(self.original_df[self.original_df['Status'] == 'Active'])

            # Populate Selectors
            self.year_cb.config(values=sorted(self.original_df[self.date_column].dt.year.unique(), reverse=True), state="readonly")
            self.year_cb.set(2026)
            self.month_cb.config(values=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], state="readonly")
            self.month_cb.set("May")
            self.day_cb.config(values=["Any"] + sorted(self.original_df[self.date_column].dt.day.unique().astype(str).tolist(), key=int), state="readonly")
            self.day_cb.set(15)

            self.all_btn.config(state="normal")
            self.match_btn.config(state="normal")
            self.filter_btn.config(state="normal")
            self.show_all()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file: {e}")

    def upload_mapping(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
        if not path: return
        try:
            df2 = pd.read_excel(path)
            ro_col_2 = next((c for c in df2.columns if "RO" in str(c).upper()), df2.columns[0])
            
            df1_match = self.original_df.copy()
            df1_match['RO Code'] = df1_match['RO Code'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df2_match = pd.DataFrame({'RO Code': df2[ro_col_2].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)})
            
            merged = pd.merge(df2_match, df1_match, on='RO Code', how='left')
            merged['Divisional Office Name'] = merged['Divisional Office Name'].fillna('Unknown')
            
            div_summary = []
            for div, group in merged.groupby('Divisional Office Name'):
                active_in_div = len(group[group['Status'] == 'Active'])
                share = (active_in_div / self.total_active_master * 100) if self.total_active_master > 0 else 0
                div_summary.append(f"{div}: {active_in_div} Active ({share:.2f}%)")

            self.calc_info_label.config(text=" | ".join(div_summary), wraplength=1200)
            self.update_dashboard(merged.sort_values('Divisional Office Name'))
            
        except Exception as e:
            messagebox.showerror("Mapping error", str(e))

    def download_table(self):
        if self.current_df is None or self.current_df.empty:
            messagebox.showwarning("Warning", "No data available to download.")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 filetypes=[("Excel file", "*.xlsx")],
                                                 title="Save Table Data")
        if not save_path: return
        
        try:
            # Create a clean version for export (formatting dates)
            export_df = self.current_df.copy()
            if self.date_column in export_df.columns:
                export_df[self.date_column] = export_df[self.date_column].dt.strftime('%d-%b-%Y')
            
            export_df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Data successfully saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Download Error", f"Failed to save file: {e}")

    def show_all(self):
        self.update_dashboard(self.original_df)

    def apply_filter(self):
        month = datetime.datetime.strptime(self.month_cb.get(), "%B").month
        year = int(self.year_cb.get())
        day = self.day_cb.get()
        mask = (self.original_df[self.date_column].dt.month == month) & (self.original_df[self.date_column].dt.year == year)
        if day != "Any": mask &= (self.original_df[self.date_column].dt.day == int(day))
        self.update_dashboard(self.original_df[mask])

    def update_dashboard(self, df):
        self.current_df = df # Store current view for downloading
        self.download_btn.config(state="normal")
        
        for w in self.charts_frame.winfo_children(): w.destroy()
        for w in self.table_canvas.winfo_children(): w.destroy()
        plt.close('all')

        if df.empty:
            self.summary_label.config(text="No Records Found")
            return

        self.summary_label.config(text=f"Records: {len(df)} | Active: {len(df[df['Status']=='Active'])} | Inactive: {len(df[df['Status']=='Inactive'])}")

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 5))
        fig.patch.set_facecolor('#001a33')

        div_df = df.groupby(['Divisional Office Name', 'Status']).size().unstack(fill_value=0)
        for s in ['Active', 'Inactive']:
            if s not in div_df.columns: div_df[s] = 0
        div_df[['Active', 'Inactive']].plot(kind='bar', ax=ax1, color=[self.status_colors['Active'], self.status_colors['Inactive']])
        ax1.set_facecolor('#001a33')
        ax1.tick_params(colors='white', labelsize=8)
        ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45, ha='right')

        counts = df['Status'].value_counts()
        ax2.pie(counts, labels=counts.index, autopct='%1.1f%%', colors=[self.status_colors[x] for x in counts.index], textprops={'color':"white"})

        plt.tight_layout()
        FigureCanvasTkAgg(fig, master=self.charts_frame).get_tk_widget().pack(fill=tk.X)

        cols = ['RO Code', 'Divisional Office Name', self.date_column, 'SAP Last 3 Months Sale', 'Status']
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview.Heading", background="#ff851b", foreground="black", font=('Arial', 10, 'bold'))
        style.configure("Custom.Treeview", background="#002b54", foreground="white", fieldbackground="#002b54", rowheight=30)

        tree = ttk.Treeview(self.table_canvas, columns=cols, show='headings', height=12, style="Custom.Treeview")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=200, anchor=tk.CENTER)
        
        for _, row in df.iterrows():
            vals = list(row[cols])
            vals[2] = vals[2].strftime('%d-%b-%Y') if pd.notnull(vals[2]) else "N/A"
            tree.insert('', tk.END, values=vals)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(self.table_canvas, orient="vertical", command=tree.yview)
        tree.configure(yscroll=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

if __name__ == "__main__":
    root = tk.Tk()
    app = IndianOilAnalyzerApp(root)
    root.mainloop()