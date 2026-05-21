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

        # --- Main Scrollable Container ---
        self.canvas = tk.Canvas(root, bg="#001a33", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#001a33")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Mouse wheel scrolling
        self.root.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self.setup_ui()

    def setup_ui(self):
        # 1. Header
        tk.Label(self.scrollable_frame, text="💾 Indian Oil RO Online Communicator", 
                 font=("Arial", 28, "bold"), fg="#ff851b", bg="#001a33").pack(pady=20)

        # 2. Top Buttons
        btn_frame = tk.Frame(self.scrollable_frame, bg="#001a33")
        btn_frame.pack(pady=10)

        self.upload_btn = tk.Button(btn_frame, text="1. Upload Excel File", command=self.upload_file, 
                                   bg="#ff851b", fg="white", font=("Arial", 12, "bold"), width=20)
        self.upload_btn.pack(side=tk.LEFT, padx=10)

        self.all_btn = tk.Button(btn_frame, text="Show All Data", command=self.show_all, 
                                bg="#17a2b8", fg="white", font=("Arial", 12, "bold"), width=20, state="disabled")
        self.all_btn.pack(side=tk.LEFT, padx=10)

        # 3. Filter Section
        self.filter_frame = tk.LabelFrame(self.scrollable_frame, text=" 2. Filter by Specific Date ", 
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

        self.filter_btn = tk.Button(self.filter_frame, text="Apply Date Filter", command=self.apply_filter, 
                                   bg="#28a745", fg="white", font=("Arial", 10, "bold"), state="disabled")
        self.filter_btn.grid(row=0, column=6, padx=20)

        # 4. Summary Stats
        self.summary_label = tk.Label(self.scrollable_frame, text="", font=("Arial", 18, "bold"), fg="white", bg="#001a33")
        self.summary_label.pack(pady=15)

        # 5. Charts Container
        self.charts_frame = tk.Frame(self.scrollable_frame, bg="#001a33")
        self.charts_frame.pack(fill=tk.X, padx=20)

        # 6. Styled Table Section (Matching your image)
        self.table_main_container = tk.Frame(self.scrollable_frame, bg="#001a33", bd=2, relief=tk.SUNKEN)
        self.table_main_container.pack(fill=tk.BOTH, expand=True, padx=40, pady=(30, 50))
        
        self.table_canvas = tk.Frame(self.table_main_container, bg="#001a33")
        self.table_canvas.pack(fill=tk.BOTH, expand=True)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls"), ("CSV", "*.csv")])
        if not path: return
        try:
            df = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
            df['Status'] = df['SAP Last 3 Months Sale'].apply(categorize_pump_status)
            df[self.date_column] = pd.to_datetime(df[self.date_column], errors='coerce')
            self.original_df = df.dropna(subset=[self.date_column])

            # Populate Selectors
            self.year_cb.config(values=sorted(self.original_df[self.date_column].dt.year.unique(), reverse=True), state="readonly")
            self.year_cb.set(2026)
            self.month_cb.config(values=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], state="readonly")
            self.month_cb.set("June")
            self.day_cb.config(values=["Any"] + sorted(self.original_df[self.date_column].dt.day.unique().astype(str).tolist(), key=int), state="readonly")
            self.day_cb.set(15)

            self.all_btn.config(state="normal")
            self.filter_btn.config(state="normal")
            self.show_all()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file: {e}")

    def show_all(self):
        self.update_dashboard(self.original_df)

    def apply_filter(self):
        month = datetime.datetime.strptime(self.month_cb.get(), "%B").month
        year = int(self.year_cb.get())
        day = self.day_cb.get()

        mask = (self.original_df[self.date_column].dt.month == month) & (self.original_df[self.date_column].dt.year == year)
        if day != "Any":
            mask &= (self.original_df[self.date_column].dt.day == int(day))
        
        filtered = self.original_df[mask]
        self.update_dashboard(filtered)

    def update_dashboard(self, df):
        for w in self.charts_frame.winfo_children(): w.destroy()
        for w in self.table_canvas.winfo_children(): w.destroy()

        if df.empty:
            self.summary_label.config(text="No Records Found for this selection")
            return

        # Summary Update
        self.summary_label.config(text=f"Showing {len(df)} Records | Active: {len(df[df['Status']=='Active'])} | Inactive: {len(df[df['Status']=='Inactive'])}")

        # Charts Update
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 5))
        fig.patch.set_facecolor('#001a33')

        # Bar Chart
        div_df = df.groupby(['Divisional Office Name', 'Status']).size().unstack(fill_value=0)
        for s in ['Active', 'Inactive']:
            if s not in div_df.columns: div_df[s] = 0
        div_df[['Active', 'Inactive']].plot(kind='bar', ax=ax1, color=[self.status_colors['Active'], self.status_colors['Inactive']])
        ax1.set_title("Status by Division", color="#ff851b", fontweight="bold", fontsize=14)
        ax1.set_facecolor('#001a33')
        ax1.tick_params(colors='white', labelsize=9)
        ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45, ha='right')

        # Pie Chart
        counts = df['Status'].value_counts()
        ax2.pie(counts, labels=counts.index, autopct='%1.1f%%', colors=[self.status_colors[x] for x in counts.index], 
                textprops={'color':"white", 'weight':'bold', 'fontsize':11})
        ax2.set_title("Overall Market Share", color="#ff851b", fontweight="bold", fontsize=14)

        plt.tight_layout()
        fig.subplots_adjust(bottom=0.35)
        FigureCanvasTkAgg(fig, master=self.charts_frame).get_tk_widget().pack(fill=tk.X)

        # --- ORANGE & BLUE STYLED TABLE (Matching the Image) ---
        cols = ['RO Code', 'Divisional Office Name', self.date_column, 'SAP Last 3 Months Sale', 'Status']
        
        style = ttk.Style()
        style.theme_use("clam")
        # Header Style: Orange Background, Black Text
        style.configure("Custom.Treeview.Heading", background="#ff851b", foreground="black", font=('Arial', 10, 'bold'), borderwidth=1)
        # Body Style: Dark Blue Background, White Text
        style.configure("Custom.Treeview", background="#002b54", foreground="white", fieldbackground="#002b54", 
                        rowheight=30, font=('Arial', 10), borderwidth=0)
        style.map("Custom.Treeview", background=[('selected', '#17a2b8')])

        tree = ttk.Treeview(self.table_canvas, columns=cols, show='headings', height=12, style="Custom.Treeview")
        
        for c in cols:
            tree.heading(c, text=c, anchor=tk.CENTER)
            tree.column(c, width=200, anchor=tk.CENTER)
        
        for _, row in df[cols].iterrows():
            vals = list(row)
            # Format numbers to look like the image (e.g., .0)
            vals[0] = f"{float(vals[0])}"
            vals[2] = vals[2].strftime('%d-%b-%Y')
            vals[3] = f"{float(vals[3])}"
            tree.insert('', tk.END, values=vals)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Table Scrollbar
        sb = ttk.Scrollbar(self.table_canvas, orient="vertical", command=tree.yview)
        tree.configure(yscroll=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

if __name__ == "__main__":
    root = tk.Tk()
    app = IndianOilAnalyzerApp(root)
    root.mainloop()