import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class IndianOilComplaintAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("INDIAN OIL COMPLAINT LIST")
        self.root.geometry("1500x950")
        self.root.configure(bg="#001a33") 

        self.primary_orange = "#ff851b" 
        self.sap_df = None  
        self.current_df = None 
        
        self.setup_styles()
        self.setup_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # --- SUMMARY TABLE STYLE (Blue Theme) ---
        style.configure("Summary.Treeview", 
                        background="#002b54", 
                        foreground="white", 
                        fieldbackground="#002b54", 
                        rowheight=28, 
                        font=("Arial", 10))
        style.configure("Summary.Treeview.Heading", 
                        background="#004080", 
                        foreground="white", 
                        font=("Arial", 10, "bold"))
        
        # --- DETAIL TABLE STYLE (Orange Trim) ---
        style.configure("Detail.Treeview", 
                        background="#002b54", 
                        foreground="white", 
                        fieldbackground="#002b54", 
                        rowheight=28, 
                        font=("Arial", 10))
        style.configure("Detail.Treeview.Heading", 
                        background=self.primary_orange, 
                        foreground="black", 
                        font=("Arial", 10, "bold"))

    def setup_ui(self):
        # Header Text
        tk.Label(self.root, text="📋 INDIAN OIL COMPLAINT LIST", 
                 font=("Arial", 28, "bold"), fg=self.primary_orange, bg="#001a33").pack(pady=15)

        # Buttons with Image-Matching Colors
        btn_frame = tk.Frame(self.root, bg="#001a33")
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="1. Load SAP Sales Data", command=self.load_sap_data, 
                  bg="#007bff", fg="white", font=("Arial", 10, "bold"), width=28, height=2, relief="flat").pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame, text="2. Upload Complaint File", command=self.upload_complaint_file, 
                  bg=self.primary_orange, fg="white", font=("Arial", 10, "bold"), width=28, height=2, relief="flat").pack(side=tk.LEFT, padx=10)

        self.download_btn = tk.Button(btn_frame, text="3. Download Records (.xlsx)", command=self.download_data, 
                                     bg="#28a745", fg="white", font=("Arial", 10, "bold"), width=28, height=2, relief="flat", state="disabled")
        self.download_btn.pack(side=tk.LEFT, padx=10)

        # Main Layout Container
        main_container = tk.Frame(self.root, bg="#001a33")
        main_container.pack(fill=tk.BOTH, expand=True, padx=40, pady=10)

        # 1. Summary Report (Upper Half)
        summary_container = tk.LabelFrame(main_container, text=" DIVISION & VENDOR ANALYSIS ", 
                                          fg="#17a2b8", bg="#001a33", font=("Arial", 12, "bold"), padx=10, pady=10)
        summary_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.summary_tree = ttk.Treeview(summary_container, columns=("Vendor", "Active", "Inactive", "Total"), 
                                         show='headings', style="Summary.Treeview")
        for col in ["Vendor", "Active", "Inactive", "Total"]:
            self.summary_tree.heading(col, text=col.upper())
            self.summary_tree.column(col, anchor="center")
        self.summary_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        s_scroll = ttk.Scrollbar(summary_container, orient="vertical", command=self.summary_tree.yview)
        self.summary_tree.configure(yscrollcommand=s_scroll.set)
        s_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 2. Detailed Log (Lower Half - ORANGE TRIM)
        detail_container = tk.LabelFrame(main_container, text=" DETAILED COMPLAINT LOG ", 
                                         fg="white", bg="#001a33", font=("Arial", 12, "bold"), padx=10, pady=10)
        detail_container.pack(fill=tk.BOTH, expand=True)

        self.detail_tree = ttk.Treeview(detail_container, show='headings', style="Detail.Treeview")
        self.detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        d_scroll_v = ttk.Scrollbar(detail_container, orient="vertical", command=self.detail_tree.yview)
        d_scroll_h = ttk.Scrollbar(self.root, orient="horizontal", command=self.detail_tree.xview)
        self.detail_tree.configure(yscrollcommand=d_scroll_v.set, xscrollcommand=d_scroll_h.set)
        
        d_scroll_v.pack(side=tk.RIGHT, fill=tk.Y)
        d_scroll_h.pack(fill=tk.X, padx=40, pady=(0, 10))

    def clean_ro_code(self, series):
        return series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    def load_sap_data(self):
        path = filedialog.askopenfilename()
        if not path: return
        try:
            df = pd.read_excel(path)
            ro_col = next(c for c in df.columns if "RO CODE" in c.upper())
            stat_col = next(c for c in df.columns if "SAP STATUS" in c.upper())
            self.sap_df = df[[ro_col, stat_col]].copy()
            self.sap_df.columns = ['RO_MATCH', 'SAP_STATUS_VAL']
            self.sap_df['RO_MATCH'] = self.clean_ro_code(self.sap_df['RO_MATCH'])
            messagebox.showinfo("Success", "SAP Data Loaded.")
        except Exception as e:
            messagebox.showerror("Error", f"SAP Load Error: {e}")

    def upload_complaint_file(self):
        if self.sap_df is None:
            messagebox.showwarning("Warning", "Load SAP Data first!")
            return
        path = filedialog.askopenfilename()
        if not path: return
        try:
            df = pd.read_excel(path, skiprows=2)
            do_col = next(c for c in df.columns if "DO NAME" in c.upper())
            vend_col = next(c for c in df.columns if "AUTOMATION VENDOR" in c.upper())
            ro_col = next(c for c in df.columns if "RO CODE" in c.upper())

            df['RO_MATCH'] = self.clean_ro_code(df[ro_col])
            merged = pd.merge(df, self.sap_df, on='RO_MATCH', how='left')
            merged['SAP_STATUS_VAL'] = merged['SAP_STATUS_VAL'].fillna('Unknown')
            
            self.current_df = merged
            self.refresh_ui(do_col, vend_col)
            self.download_btn.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Processing Error: {e}")

    def refresh_ui(self, do_col, vend_col):
        # Update Summary
        for item in self.summary_tree.get_children(): self.summary_tree.delete(item)
        grouped = self.current_df.groupby([do_col, vend_col, 'SAP_STATUS_VAL']).size().unstack(fill_value=0)
        
        for division in sorted(self.current_df[do_col].unique()):
            self.summary_tree.insert("", tk.END, values=(f"▶ {division}", "", "", ""), tags=('div_header',))
            div_data = grouped.loc[division] if division in grouped.index else pd.DataFrame()
            d_act, d_inact = 0, 0
            for vendor in div_data.index:
                act = div_data.loc[vendor, 'Active'] if 'Active' in div_data.columns else 0
                inact = div_data.loc[vendor, 'Inactive'] if 'Inactive' in div_data.columns else 0
                self.summary_tree.insert("", tk.END, values=(f"   {vendor}", act, inact, act+inact))
                d_act += act; d_inact += inact
            self.summary_tree.insert("", tk.END, values=("   TOTAL", d_act, d_inact, d_act+d_inact), tags=('subtotal',))
        
        self.summary_tree.tag_configure('div_header', background='#004080', foreground=self.primary_orange, font=('Arial', 10, 'bold'))
        self.summary_tree.tag_configure('subtotal', background='#001f3f', foreground="#28a745")

        # Update Details
        detail_cols = ['COMP ID', 'RO CODE', 'RO NAME', vend_col, 'FLAG', 'AGE SINCE COMPLAINT CREATED', 'SAP_STATUS_VAL']
        self.detail_tree["columns"] = detail_cols
        for c in detail_cols:
            self.detail_tree.heading(c, text=c)
            self.detail_tree.column(c, width=180, anchor="center")
        
        for item in self.detail_tree.get_children(): self.detail_tree.delete(item)
        for _, row in self.current_df[detail_cols].iterrows():
            self.detail_tree.insert("", tk.END, values=list(row))

    def download_data(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if save_path:
            # STRICT EXPORT (Based on your Screenshot)
            export_cols = ['COMP ID', 'RO CODE', 'RO NAME', 'DO NAME', 'COMPLAINT DATETIME', 
                           'AUTOMATION VENDOR', 'ISSUE TYPE', 'ISSUE SUB TYPE', 
                           'ADDITIONAL_DETAILS', 'FLAG', 'AGE SINCE COMPLAINT CREATED']
            
            final_cols = [c for c in export_cols if c in self.current_df.columns]
            self.current_df[final_cols].to_excel(save_path, index=False)
            messagebox.showinfo("Success", "Excel Exported Successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = IndianOilComplaintAnalyzer(root)
    root.mainloop()