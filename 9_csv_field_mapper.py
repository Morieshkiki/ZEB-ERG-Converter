"""
CSV Field Mapper
A Python tkinter application that allows users to:
1. Load a CSV file
2. View column headers from the CSV
3. Map CSV columns to predefined target fields (user-editable)
4. Calculate/transform data based on mapping rules
5. Save the data to a Microsoft Access database file (.accdb) or CSV

Note: Multiple target fields CAN be mapped to the same CSV column.

Requirements:
    pip install pandas

Author: Claude
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import csv
import os
import json
import re
import threading
import time
from pathlib import Path

# Check for pandas
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


class ProgressDialog:
    """A progress dialog with elapsed time display."""
    
    def __init__(self, parent, title="Processing"):
        self.parent = parent
        self.cancelled = False
        self.start_time = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x180")
        self.dialog.configure(bg="#1a1a2e")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 200
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 90
        self.dialog.geometry(f"+{x}+{y}")
        
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self.title_label = tk.Label(
            self.dialog, text=title, fg="white", bg="#1a1a2e",
            font=("Segoe UI", 14, "bold")
        )
        self.title_label.pack(pady=(20, 10))
        
        self.status_label = tk.Label(
            self.dialog, text="Initializing...", fg="#888888", bg="#1a1a2e",
            font=("Segoe UI", 10)
        )
        self.status_label.pack(pady=(0, 10))
        
        style = ttk.Style()
        style.configure("Custom.Horizontal.TProgressbar", 
                       troughcolor='#252542', background='#6366f1')
        
        self.progress = ttk.Progressbar(
            self.dialog, length=350, mode='determinate',
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=10)
        
        self.time_label = tk.Label(
            self.dialog, text="Elapsed: 0.0s", fg="#a855f7", bg="#1a1a2e",
            font=("Segoe UI", 10)
        )
        self.time_label.pack(pady=(5, 10))
        
        self.row_label = tk.Label(
            self.dialog, text="", fg="#22c55e", bg="#1a1a2e",
            font=("Segoe UI", 9)
        )
        self.row_label.pack(pady=(0, 10))
        
    def start(self):
        self.start_time = time.time()
        self._update_time()
        
    def _update_time(self):
        if self.start_time and not self.cancelled:
            elapsed = time.time() - self.start_time
            self.time_label.config(text=f"Elapsed: {elapsed:.1f}s")
            self.dialog.after(100, self._update_time)
            
    def update(self, current, total, status_text=""):
        percent = (current / total * 100) if total > 0 else 0
        self.progress['value'] = percent
        if status_text:
            self.status_label.config(text=status_text)
        self.row_label.config(text=f"Row {current:,} of {total:,}")
        self.dialog.update()
        
    def close(self):
        self.cancelled = True
        self.dialog.destroy()


class EditTargetFieldsDialog:
    """Dialog for editing target fields."""
    
    def __init__(self, parent, current_fields):
        self.parent = parent
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Edit Target Fields")
        self.dialog.geometry("500x600")
        self.dialog.configure(bg="#1a1a2e")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 250
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 300
        self.dialog.geometry(f"+{x}+{y}")
        
        tk.Label(
            self.dialog, text="Edit Target Fields", fg="white", bg="#1a1a2e",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(15, 5))
        
        tk.Label(
            self.dialog, text="One field name per line", fg="#888888", bg="#1a1a2e",
            font=("Segoe UI", 9)
        ).pack(pady=(0, 10))
        
        text_frame = tk.Frame(self.dialog, bg="#252542", padx=2, pady=2)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.text_area = tk.Text(
            text_frame, bg="#1e1e38", fg="#e0e0e0", font=("Consolas", 10),
            insertbackground="white", selectbackground="#6366f1"
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)
        self.text_area.insert("1.0", "\n".join(current_fields))
        
        btn_frame = tk.Frame(self.dialog, bg="#1a1a2e")
        btn_frame.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Button(
            btn_frame, text="Cancel", command=self.cancel,
            bg="#374151", fg="white", font=("Segoe UI", 10),
            padx=20, pady=8, relief=tk.FLAT, cursor="hand2"
        ).pack(side=tk.LEFT)
        
        tk.Button(
            btn_frame, text="Save Fields", command=self.save,
            bg="#22c55e", fg="white", font=("Segoe UI", 10, "bold"),
            padx=20, pady=8, relief=tk.FLAT, cursor="hand2"
        ).pack(side=tk.RIGHT)
        
        self.dialog.wait_window()
        
    def cancel(self):
        self.result = None
        self.dialog.destroy()
        
    def save(self):
        text = self.text_area.get("1.0", tk.END)
        fields = [f.strip() for f in text.strip().split("\n") if f.strip()]
        if not fields:
            messagebox.showwarning("Warning", "At least one field is required.")
            return
        self.result = fields
        self.dialog.destroy()


class CSVFieldMapper:
    """Main application class for CSV field mapping."""
    
    DEFAULT_TARGET_FIELDS = [
        "ID", "ZENDI", "KLASSE", "NUMMER", "BUCHSTABE", "LAGE", "FS", "VNK", "NNK",
        "VST", "BST", "VKM", "BKM", "FSANZAHL", "FBANZAHL", "OD_FS", "BAULAST",
        "RADWEG_FLA", "RAD", "Bauw_3", "BREITE", "RIS", "RISK", "FLI", "FLIK",
        "SUB", "VER", "AUS", "SEN", "HEB", "RISG", "FLIG", "RSF", "KUN", "GEF",
        "G", "BAUW_PR", "DATUM_3", "Uhr_3", "ZWRIS", "ZWRISK", "ZWFLI", "ZWFLIK",
        "ZWSUB", "ZWVER", "ZWAUS", "ZWSEN", "ZWHEB", "ZWRISG", "ZWFLIG", "ZWKUN",
        "TWGEB", "TWSUB", "GW", "NIC_L", "NIC_R", "NIC", "ZWNIC", "FUNKTION",
        "BAUW", "DATUM_1A", "UHRZEIT_1A", "VM_1A", "AUN", "ZWAUN", "PGR_AVG",
        "PGR_MAX", "ZWPGR", "SBL", "DBL", "W", "LWI_FS", "RML_LWI_FS", "LWI_OD",
        "RML_LWI_OD", "ZWLWI", "S03", "S10", "S30", "LN", "K", "DATUM_1B",
        "UHRZEIT_1B", "VM_1B", "MSPTR", "MSPTL", "MSPT", "ZWSPT", "MSPHR", "MSPHL",
        "MSPH", "ZWSPH", "SSPTR", "SSPTL", "SSPHR", "SSPHL", "QN", "DATUM_2",
        "UHRZEIT_2", "VMIN_2", "GRI_40", "GRI_60", "GRI_80", "ZWGRI", "UHRZEIT_3",
        "VM_3", "RISS", "ZWRISS", "EFLI", "AFLI", "ZWAFLI", "ONA", "BIN", "RSFA",
        "ZWRSFA", "LQRL", "ZWLQRL", "LQRP", "ZWLQRP", "LQR", "ZWLQR", "EABF",
        "ZWEABF", "EABP", "ZWEABP", "EAB", "ZWEAB", "KASL", "ZWKASL", "KASP",
        "ZWKASP", "ZWKAS", "RSFB", "ZWRSFB", "NTR", "FUF", "BTE", "TWUM", "ZK",
        "MESSJAHR", "ZWAUN_15", "ZWLWI_15", "ZWDBL_15", "ZWSBL_15", "ZWBPL_15",
        "ZWSPT_15", "ZWSPH_15", "ZWGRI_15", "ZWRISS_15", "ZWFLI_15", "ZWAFLI_15",
        "ZWRISG_15", "ZWLQRL_15", "ZWLQRP_15", "ZWLQR_15", "ZWRSFB_15", "ZWRSFA_15",
        "TWE_15", "TWN_15", "TWEQLQ_15", "TWRIO_15", "GEB_15", "SUB_15", "GW_15",
        "OFS", "IRI", "ZWPGR_AVG", "ZWPGR_MAX", "ZWEFLI", "ZWONA", "ZWBIN", "ZWOFS",
        "ZWSCH", "ZWBORD", "ZWWURZ", "ZWRSF", "GEB",
    ]
    
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Field Mapper")
        self.root.geometry("1100x750")
        self.root.configure(bg="#1a1a2e")
        
        self.TARGET_FIELDS = self.DEFAULT_TARGET_FIELDS.copy()
        self.csv_headers = []
        self.csv_data = []
        self.mappings = {}
        self.csv_file_path = None
        self.calculated_data = None
        self.selected_csv_field = tk.StringVar(value="")
        self.csv_delimiter = ','  # Will be auto-detected
        
        self.setup_styles()
        self.create_widgets()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#1a1a2e")
        style.configure("TLabel", background="#1a1a2e", foreground="#e0e0e0", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground="#a855f7")
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#6366f1")
        style.configure("Status.TLabel", font=("Segoe UI", 9), foreground="#888888")
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(main_frame, text="CSV Field Mapper", style="Title.TLabel")
        title_label.pack(pady=(0, 5))
        
        subtitle_label = ttk.Label(main_frame, text="Map CSV columns to target fields → Calculate → Export to Access/CSV", style="Status.TLabel")
        subtitle_label.pack(pady=(0, 20))
        
        self.create_button_bar(main_frame)
        
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=0)
        content_frame.columnconfigure(2, weight=1)
        content_frame.rowconfigure(0, weight=1)
        
        self.create_target_panel(content_frame)
        self.create_center_panel(content_frame)
        self.create_csv_panel(content_frame)
        self.create_status_bar(main_frame)
        
    def create_button_bar(self, parent):
        # First row: Edit Target Fields (big) + Load CSV File (big) + Calculate (big, right side)
        button_frame1 = ttk.Frame(parent)
        button_frame1.pack(fill=tk.X, pady=(0, 8))
        
        self.edit_fields_btn = tk.Button(
            button_frame1, text="✏️ Edit Target Fields", command=self.edit_target_fields,
            bg="#f59e0b", fg="white", font=("Segoe UI", 11, "bold"),
            padx=25, pady=10, relief=tk.FLAT, cursor="hand2"
        )
        self.edit_fields_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.load_btn = tk.Button(
            button_frame1, text="📂 Load CSV File", command=self.load_csv,
            bg="#6366f1", fg="white", font=("Segoe UI", 11, "bold"),
            padx=25, pady=10, relief=tk.FLAT, cursor="hand2"
        )
        self.load_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.calculate_btn = tk.Button(
            button_frame1, text="🔢 Calculate", command=self.calculate_data,
            bg="#ec4899", fg="white", font=("Segoe UI", 11, "bold"),
            padx=30, pady=10, relief=tk.FLAT, cursor="hand2", state=tk.DISABLED
        )
        self.calculate_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.calc_status_label = tk.Label(
            button_frame1, text="", fg="#ec4899", bg="#1a1a2e",
            font=("Segoe UI", 9, "italic")
        )
        self.calc_status_label.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Second row: Auto-Map buttons (smaller) + Clear All (small)
        button_frame2 = ttk.Frame(parent)
        button_frame2.pack(fill=tk.X, pady=(0, 15))
        
        self.automap_btn = tk.Button(
            button_frame2, text="⚡ Auto-Map by Position", command=self.auto_map,
            bg="#8b5cf6", fg="white", font=("Segoe UI", 9),
            padx=12, pady=6, relief=tk.FLAT, cursor="hand2", state=tk.DISABLED
        )
        self.automap_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        self.smart_automap_btn = tk.Button(
            button_frame2, text="🎯 Smart Auto-Map", command=self.smart_auto_map,
            bg="#0891b2", fg="white", font=("Segoe UI", 9),
            padx=12, pady=6, relief=tk.FLAT, cursor="hand2", state=tk.DISABLED
        )
        self.smart_automap_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        self.clear_btn = tk.Button(
            button_frame2, text="🗑 Clear All", command=self.clear_mappings,
            bg="#374151", fg="#ef4444", font=("Segoe UI", 9),
            padx=10, pady=6, relief=tk.FLAT, cursor="hand2"
        )
        self.clear_btn.pack(side=tk.LEFT)
        
    def create_target_panel(self, parent):
        container = tk.Frame(parent, bg="#252542", padx=2, pady=2)
        container.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        inner_frame = tk.Frame(container, bg="#1a1a2e")
        inner_frame.pack(fill=tk.BOTH, expand=True)
        
        header_frame = tk.Frame(inner_frame, bg="#1a1a2e")
        header_frame.pack(fill=tk.X, padx=15, pady=15)
        
        tk.Label(header_frame, text="●", fg="#6366f1", bg="#1a1a2e", font=("Segoe UI", 12)).pack(side=tk.LEFT)
        tk.Label(header_frame, text="Target Fields (Access)", fg="white", bg="#1a1a2e", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(8, 0))
        
        self.target_count_label = tk.Label(header_frame, text=f"0/{len(self.TARGET_FIELDS)} mapped", fg="#666666", bg="#1a1a2e", font=("Segoe UI", 9))
        self.target_count_label.pack(side=tk.RIGHT)
        
        tk.Frame(inner_frame, height=1, bg="#333355").pack(fill=tk.X, padx=15)
        
        list_container = tk.Frame(inner_frame, bg="#1a1a2e")
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.target_canvas = tk.Canvas(list_container, bg="#1a1a2e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.target_canvas.yview)
        
        self.target_list_frame = tk.Frame(self.target_canvas, bg="#1a1a2e")
        
        self.target_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.target_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.target_canvas_frame = self.target_canvas.create_window((0, 0), window=self.target_list_frame, anchor="nw")
        
        self.target_list_frame.bind("<Configure>", lambda e: self.target_canvas.configure(scrollregion=self.target_canvas.bbox("all")))
        self.target_canvas.bind("<Configure>", lambda e: self.target_canvas.itemconfig(self.target_canvas_frame, width=e.width))
        
        # Bind mouse wheel for target panel
        self.target_canvas.bind("<Enter>", lambda e: self._bind_mousewheel(self.target_canvas))
        self.target_canvas.bind("<Leave>", lambda e: self._unbind_mousewheel())
        
        self.target_widgets = {}
        for i, field in enumerate(self.TARGET_FIELDS):
            self.create_target_row(i, field)
    
    def _bind_mousewheel(self, canvas):
        """Bind mousewheel to specific canvas."""
        self.active_canvas = canvas
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.root.bind_all("<Button-5>", self._on_mousewheel_linux)
    
    def _unbind_mousewheel(self):
        """Unbind mousewheel."""
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")
        self.active_canvas = None
    
    def _on_mousewheel(self, event):
        """Handle mousewheel scroll (Windows/MacOS)."""
        if hasattr(self, 'active_canvas') and self.active_canvas:
            self.active_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _on_mousewheel_linux(self, event):
        """Handle mousewheel scroll (Linux)."""
        if hasattr(self, 'active_canvas') and self.active_canvas:
            if event.num == 4:
                self.active_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.active_canvas.yview_scroll(1, "units")
            
    def create_target_row(self, index, field_name):
        row_frame = tk.Frame(self.target_list_frame, bg="#1e1e38", padx=10, pady=8)
        row_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(row_frame, text=f"{index+1:02d}", fg="#666666", bg="#1e1e38", font=("Consolas", 9)).pack(side=tk.LEFT)
        
        name_label = tk.Label(row_frame, text=field_name, fg="#cccccc", bg="#1e1e38", font=("Segoe UI", 10), width=14, anchor="w")
        name_label.pack(side=tk.LEFT, padx=(10, 10))
        
        optional_label = tk.Label(row_frame, text="(optional)", fg="#555555", bg="#1e1e38", font=("Segoe UI", 8, "italic"))
        optional_label.pack(side=tk.LEFT, padx=(0, 10))
        
        mapping_var = tk.StringVar(value="")
        mapping_combo = ttk.Combobox(row_frame, textvariable=mapping_var, state="readonly", width=20, font=("Segoe UI", 9))
        mapping_combo.pack(side=tk.LEFT, padx=(0, 10))
        mapping_combo.bind("<<ComboboxSelected>>", lambda e, f=field_name, v=mapping_var: self.on_mapping_selected(f, v))
        
        clear_btn = tk.Button(
            row_frame, text="×", command=lambda f=field_name: self.clear_single_mapping(f),
            bg="#1e1e38", fg="#ef4444", font=("Segoe UI", 12, "bold"),
            relief=tk.FLAT, cursor="hand2", padx=5, pady=0
        )
        clear_btn.pack(side=tk.RIGHT)
        
        self.target_widgets[field_name] = {
            "frame": row_frame, "combo": mapping_combo, "var": mapping_var,
            "name_label": name_label, "optional_label": optional_label, "clear_btn": clear_btn
        }
        
    def create_center_panel(self, parent):
        center_frame = tk.Frame(parent, bg="#1a1a2e", width=80)
        center_frame.grid(row=0, column=1, sticky="ns", padx=10)
        center_frame.grid_propagate(False)
        
        tk.Frame(center_frame, bg="#1a1a2e", height=150).pack()
        
        arrow_canvas = tk.Canvas(center_frame, width=60, height=60, bg="#1a1a2e", highlightthickness=0)
        arrow_canvas.pack()
        arrow_canvas.create_oval(5, 5, 55, 55, fill="#6366f1", outline="#8b5cf6", width=2)
        arrow_canvas.create_text(30, 30, text="←", fill="white", font=("Segoe UI", 20, "bold"))
        
        tk.Label(center_frame, text="Select from\ndropdowns\nto map", fg="#666666", bg="#1a1a2e", font=("Segoe UI", 8), justify=tk.CENTER).pack(pady=10)
        tk.Label(center_frame, text="Empty fields\nwill be\nexcluded", fg="#888800", bg="#1a1a2e", font=("Segoe UI", 8), justify=tk.CENTER).pack(pady=10)
        tk.Label(center_frame, text="Same column\ncan map to\nmultiple fields", fg="#0891b2", bg="#1a1a2e", font=("Segoe UI", 8), justify=tk.CENTER).pack(pady=10)
        
    def create_csv_panel(self, parent):
        container = tk.Frame(parent, bg="#252542", padx=2, pady=2)
        container.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        
        inner_frame = tk.Frame(container, bg="#1a1a2e")
        inner_frame.pack(fill=tk.BOTH, expand=True)
        
        header_frame = tk.Frame(inner_frame, bg="#1a1a2e")
        header_frame.pack(fill=tk.X, padx=15, pady=15)
        
        tk.Label(header_frame, text="●", fg="#a855f7", bg="#1a1a2e", font=("Segoe UI", 12)).pack(side=tk.LEFT)
        tk.Label(header_frame, text="CSV Columns", fg="white", bg="#1a1a2e", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(8, 0))
        
        self.csv_file_label = tk.Label(header_frame, text="No file loaded", fg="#666666", bg="#1a1a2e", font=("Segoe UI", 9))
        self.csv_file_label.pack(side=tk.RIGHT)
        
        tk.Frame(inner_frame, height=1, bg="#333355").pack(fill=tk.X, padx=15)
        
        list_container = tk.Frame(inner_frame, bg="#1a1a2e")
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.csv_canvas = tk.Canvas(list_container, bg="#1a1a2e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.csv_canvas.yview)
        
        self.csv_list_frame = tk.Frame(self.csv_canvas, bg="#1a1a2e")
        
        self.csv_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.csv_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.csv_canvas_frame = self.csv_canvas.create_window((0, 0), window=self.csv_list_frame, anchor="nw")
        
        self.csv_list_frame.bind("<Configure>", lambda e: self.csv_canvas.configure(scrollregion=self.csv_canvas.bbox("all")))
        self.csv_canvas.bind("<Configure>", lambda e: self.csv_canvas.itemconfig(self.csv_canvas_frame, width=e.width))
        
        # Bind mouse wheel for CSV panel
        self.csv_canvas.bind("<Enter>", lambda e: self._bind_mousewheel(self.csv_canvas))
        self.csv_canvas.bind("<Leave>", lambda e: self._unbind_mousewheel())
        
        self.csv_placeholder = tk.Label(self.csv_list_frame, text="Load a CSV file to see columns here", fg="#555555", bg="#1a1a2e", font=("Segoe UI", 10, "italic"))
        self.csv_placeholder.pack(pady=50)
        
        self.csv_widgets = {}
        
    def create_csv_row(self, index, column_name):
        row_frame = tk.Frame(self.csv_list_frame, bg="#2d1f4e", padx=10, pady=8)
        row_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(row_frame, text=f"{index+1:02d}", fg="#666666", bg="#2d1f4e", font=("Consolas", 9)).pack(side=tk.LEFT)
        
        name_label = tk.Label(row_frame, text=column_name, fg="#d8b4fe", bg="#2d1f4e", font=("Segoe UI", 10), anchor="w")
        name_label.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
        
        mapped_label = tk.Label(row_frame, text="", fg="#22c55e", bg="#2d1f4e", font=("Segoe UI", 9))
        mapped_label.pack(side=tk.RIGHT)
        
        self.csv_widgets[column_name] = {"frame": row_frame, "name_label": name_label, "mapped_label": mapped_label}
        
    def create_status_bar(self, parent):
        status_frame = tk.Frame(parent, bg="#252542", padx=15, pady=10)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.mapped_count_label = tk.Label(status_frame, text="0 fields mapped", fg="#22c55e", bg="#252542", font=("Segoe UI", 9))
        self.mapped_count_label.pack(side=tk.LEFT)
        
        self.remaining_count_label = tk.Label(status_frame, text=f"{len(self.TARGET_FIELDS)} fields remaining (optional)", fg="#888888", bg="#252542", font=("Segoe UI", 9))
        self.remaining_count_label.pack(side=tk.LEFT, padx=(20, 0))
        
        self.shared_count_label = tk.Label(status_frame, text="", fg="#0891b2", bg="#252542", font=("Segoe UI", 9))
        self.shared_count_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # Save button at bottom right
        self.save_btn = tk.Button(
            status_frame, text="💾 Save to Access/CSV", command=self.save_to_access,
            bg="#22c55e", fg="white", font=("Segoe UI", 10, "bold"),
            padx=20, pady=8, relief=tk.FLAT, cursor="hand2", state=tk.DISABLED
        )
        self.save_btn.pack(side=tk.RIGHT)
        
        self.row_count_label = tk.Label(status_frame, text="", fg="#a855f7", bg="#252542", font=("Segoe UI", 9))
        self.row_count_label.pack(side=tk.RIGHT, padx=(0, 20))
        
    def edit_target_fields(self):
        # Save current mappings before editing
        old_mappings = self.mappings.copy()
        
        dialog = EditTargetFieldsDialog(self.root, self.TARGET_FIELDS)
        if dialog.result:
            self.TARGET_FIELDS = dialog.result
            self.rebuild_target_panel(old_mappings)
            self.calculated_data = None
            self.calc_status_label.config(text="")
            messagebox.showinfo("Success", f"Updated to {len(self.TARGET_FIELDS)} target fields.")
            
    def rebuild_target_panel(self, old_mappings=None):
        """Rebuild the target panel with new fields, preserving mappings where possible."""
        for widget in self.target_list_frame.winfo_children():
            widget.destroy()
        self.target_widgets.clear()
        self.mappings.clear()
        
        for i, field in enumerate(self.TARGET_FIELDS):
            self.create_target_row(i, field)
            
        # Update combos if CSV is loaded
        if self.csv_headers:
            self.update_target_combos()
            
            # Restore mappings for fields that still exist
            if old_mappings:
                for field_name, csv_col in old_mappings.items():
                    if field_name in self.target_widgets and csv_col in self.csv_headers:
                        self.mappings[field_name] = csv_col
                        self.target_widgets[field_name]["var"].set(csv_col)
                        self.update_row_style(field_name, True)
                
                self.update_csv_mapped_indicators()
            
        self.update_counts()
        
    def load_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        sample = f.read(8192)
                        f.seek(0)
                        
                        # Auto-detect delimiter (German CSVs often use semicolon)
                        first_line = sample.split('\n')[0]
                        semicolon_count = first_line.count(';')
                        comma_count = first_line.count(',')
                        tab_count = first_line.count('\t')
                        
                        # Choose delimiter based on count
                        if semicolon_count > comma_count and semicolon_count > tab_count:
                            self.csv_delimiter = ';'
                        elif tab_count > comma_count and tab_count > semicolon_count:
                            self.csv_delimiter = '\t'
                        else:
                            self.csv_delimiter = ','
                        
                        f.seek(0)
                        reader = csv.reader(f, delimiter=self.csv_delimiter)
                        
                        self.csv_headers = next(reader)
                        self.csv_headers = [h.strip().strip('\ufeff') for h in self.csv_headers]
                        self.csv_data = list(reader)
                        break
                except UnicodeDecodeError:
                    continue
            else:
                raise Exception("Could not decode file")
                
            self.csv_file_path = file_path
            self.calculated_data = None
            self.calc_status_label.config(text="")
            self.update_csv_panel()
            self.update_target_combos()
            
            self.csv_file_label.config(text=Path(file_path).name)
            self.automap_btn.config(state=tk.NORMAL)
            self.smart_automap_btn.config(state=tk.NORMAL)
            self.calculate_btn.config(state=tk.NORMAL)
            self.row_count_label.config(text=f"{len(self.csv_data)} data rows")
            
            delimiter_name = {',' : 'comma', ';': 'semicolon', '\t': 'tab'}.get(self.csv_delimiter, self.csv_delimiter)
            messagebox.showinfo("Success", f"Loaded {len(self.csv_headers)} columns and {len(self.csv_data)} rows.\nDelimiter detected: {delimiter_name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{str(e)}")
            
    def update_csv_panel(self):
        for widget in self.csv_list_frame.winfo_children():
            widget.destroy()
        self.csv_widgets.clear()
        
        for i, header in enumerate(self.csv_headers):
            self.create_csv_row(i, header)
            
    def update_target_combos(self):
        values = [""] + self.csv_headers
        for widgets in self.target_widgets.values():
            widgets["combo"]["values"] = values
            
    def on_mapping_selected(self, target_field, var):
        csv_column = var.get()
        
        if csv_column:
            self.mappings[target_field] = csv_column
            self.update_row_style(target_field, True)
        else:
            if target_field in self.mappings:
                del self.mappings[target_field]
            self.update_row_style(target_field, False)
            
        self.update_csv_mapped_indicators()
        self.update_counts()
        
        self.calculated_data = None
        self.calc_status_label.config(text="")
        self.save_btn.config(state=tk.DISABLED if not self.mappings else tk.NORMAL)
        
    def update_row_style(self, target_field, is_mapped):
        widgets = self.target_widgets[target_field]
        
        if is_mapped:
            widgets["frame"].config(bg="#1a3329")
            widgets["name_label"].config(bg="#1a3329", fg="#22c55e")
            widgets["optional_label"].config(bg="#1a3329", fg="#1a5535")
        else:
            widgets["frame"].config(bg="#1e1e38")
            widgets["name_label"].config(bg="#1e1e38", fg="#cccccc")
            widgets["optional_label"].config(bg="#1e1e38", fg="#555555")
            
    def update_csv_mapped_indicators(self):
        column_usage_count = {}
        for csv_col in self.mappings.values():
            column_usage_count[csv_col] = column_usage_count.get(csv_col, 0) + 1
        
        for col_name, widgets in self.csv_widgets.items():
            usage_count = column_usage_count.get(col_name, 0)
            
            if usage_count > 1:
                widgets["mapped_label"].config(text=f"✓ ×{usage_count}", fg="#0891b2")
                widgets["frame"].config(bg="#1f2d3d")
                widgets["name_label"].config(bg="#1f2d3d", fg="#67e8f9")
            elif usage_count == 1:
                widgets["mapped_label"].config(text="✓", fg="#22c55e")
                widgets["frame"].config(bg="#1f2d1f")
                widgets["name_label"].config(bg="#1f2d1f", fg="#888888")
            else:
                widgets["mapped_label"].config(text="")
                widgets["frame"].config(bg="#2d1f4e")
                widgets["name_label"].config(bg="#2d1f4e", fg="#d8b4fe")
                
    def update_counts(self):
        mapped = len(self.mappings)
        total = len(self.TARGET_FIELDS)
        remaining = total - mapped
        
        column_usage_count = {}
        for csv_col in self.mappings.values():
            column_usage_count[csv_col] = column_usage_count.get(csv_col, 0) + 1
        shared_columns = sum(1 for count in column_usage_count.values() if count > 1)
        
        self.mapped_count_label.config(text=f"{mapped} fields mapped")
        self.remaining_count_label.config(text=f"{remaining} fields remaining (optional)")
        self.target_count_label.config(text=f"{mapped}/{total} mapped")
        self.shared_count_label.config(text=f"{shared_columns} column(s) shared" if shared_columns > 0 else "")
        
    def auto_map(self):
        if not self.csv_headers:
            return
            
        self.clear_mappings()
        
        for i, target_field in enumerate(self.TARGET_FIELDS):
            if i < len(self.csv_headers):
                self.mappings[target_field] = self.csv_headers[i]
                self.target_widgets[target_field]["var"].set(self.csv_headers[i])
                self.update_row_style(target_field, True)
                
        self.update_csv_mapped_indicators()
        self.update_counts()
        self.calculate_btn.config(state=tk.NORMAL if self.csv_data else tk.DISABLED)
        
    def smart_auto_map(self):
        if not self.csv_headers:
            return
            
        self.clear_mappings()
        
        mapping_rules = {
            "ID": ["ID"],
            "hiline_carriageway": ["LAGE"],
            "hiline_road": ["KLASSE", "NUMMER"],
            "business_data": [
                "EFLI", "AFLI", "RISS", "ZWAUS", "ZWBIN", "ZWONA", 
                "ZWRSF", "ZWSCH", "ZWAFLI", "ZWBORD", "ZWEFLI", 
                "ZWRISS", "ZWWURZ", "GW", "GEB", "SUB"
            ],
            "hiline_section": ["VNK", "NNK"],
            "hiline_lane": ["FS"],
        }
        
        csv_header_map = {h.lower(): h for h in self.csv_headers}
        matched = 0
        
        for csv_col_key, target_fields in mapping_rules.items():
            csv_col_lower = csv_col_key.lower()
            if csv_col_lower in csv_header_map:
                actual_csv_col = csv_header_map[csv_col_lower]
                
                for target_field in target_fields:
                    if target_field in self.target_widgets:
                        self.mappings[target_field] = actual_csv_col
                        self.target_widgets[target_field]["var"].set(actual_csv_col)
                        self.update_row_style(target_field, True)
                        matched += 1
                
        self.update_csv_mapped_indicators()
        self.update_counts()
        self.calculate_btn.config(state=tk.NORMAL if self.csv_data else tk.DISABLED)
        messagebox.showinfo("Smart Auto-Map", f"Mapped {matched} fields based on predefined rules.")
        
    def clear_mappings(self):
        self.mappings.clear()
        self.calculated_data = None
        self.calc_status_label.config(text="")
        
        for field_name, widgets in self.target_widgets.items():
            widgets["var"].set("")
            self.update_row_style(field_name, False)
            
        self.update_csv_mapped_indicators()
        self.update_counts()
        self.save_btn.config(state=tk.DISABLED)
        
    def clear_single_mapping(self, target_field):
        if target_field in self.mappings:
            del self.mappings[target_field]
            
        self.target_widgets[target_field]["var"].set("")
        self.update_row_style(target_field, False)
        self.update_csv_mapped_indicators()
        self.update_counts()
        
        self.calculated_data = None
        self.calc_status_label.config(text="")

    def _extract_from_business_data(self, json_str, target_field):
        """Extract specific value from business_data JSON based on target field."""
        try:
            data = json.loads(json_str)
            
            # survey_result.tp3 fields
            survey_tp3_fields = {
                "EFLI": "efli", "AFLI": "afli", "RISS": "riss", "ONA": "ona"
            }
            
            # evaluation_result.tp3 fields
            eval_tp3_fields = {
                "ZWAUS": "zwaus", "ZWBIN": "zwbin", "ZWONA": "zwona",
                "ZWRSF": "zwrsf", "ZWSCH": "zwsch", "ZWAFLI": "zwafli",
                "ZWBORD": "zwbord", "ZWEFLI": "zwefli", "ZWRISS": "zwriss",
                "ZWWURZ": "zwwurz"
            }
            
            # evaluation_result.overall fields
            overall_fields = {
                "GW": "gw", "GEB": "geb", "SUB": "sub"
            }
            
            if target_field in survey_tp3_fields:
                key = survey_tp3_fields[target_field]
                return data.get("survey_result", {}).get("tp3", {}).get(key, "")
            elif target_field in eval_tp3_fields:
                key = eval_tp3_fields[target_field]
                return data.get("evaluation_result", {}).get("tp3", {}).get(key, "")
            elif target_field in overall_fields:
                key = overall_fields[target_field]
                return data.get("evaluation_result", {}).get("overall", {}).get(key, "")
            else:
                return ""
                
        except (json.JSONDecodeError, AttributeError, TypeError):
            return ""

    def _extract_vnk(self, hiline_section):
        """Extract VNK (first half) from hiline_section."""
        if hiline_section and len(hiline_section) >= 2:
            half_len = len(hiline_section) // 2
            return hiline_section[:half_len]
        return hiline_section
    
    def _extract_nnk(self, hiline_section):
        """Extract NNK (second half) from hiline_section."""
        if hiline_section and len(hiline_section) >= 2:
            half_len = len(hiline_section) // 2
            return hiline_section[half_len:]
        return hiline_section

    def _extract_klasse(self, hiline_road):
        """Extract KLASSE from hiline_road. Example: l0048 -> L48"""
        if not hiline_road:
            return ""
        
        letter = hiline_road[0].upper() if hiline_road else ""
        numeric_part = hiline_road[1:] if len(hiline_road) > 1 else ""
        try:
            numeric_value = int(numeric_part)
            return f"{letter}{numeric_value}"
        except ValueError:
            return f"{letter}{numeric_part}"

    def _extract_nummer(self, hiline_road):
        """Extract NUMMER from hiline_road. Example: l0048 -> 0048"""
        if not hiline_road:
            return ""
        return hiline_road[1:] if len(hiline_road) > 1 else ""

    def calculate_data(self):
        """Calculate/transform data based on mappings with progress dialog."""
        if not self.mappings:
            messagebox.showwarning("Warning", "No mappings defined. Please map some fields first.")
            return
            
        if not self.csv_data:
            messagebox.showwarning("Warning", "No CSV data loaded.")
            return
        
        progress = ProgressDialog(self.root, "Calculating Data")
        progress.start()
        
        try:
            total_rows = len(self.csv_data)
            active_mappings = {k: v for k, v in self.mappings.items() if v}
            
            header_index = {h: i for i, h in enumerate(self.csv_headers)}
            
            self.calculated_data = {field: [] for field in active_mappings.keys()}
            
            for row_idx, row in enumerate(self.csv_data):
                if row_idx % 100 == 0:
                    progress.update(row_idx, total_rows, f"Processing row {row_idx + 1}...")
                
                for target_field, csv_column in active_mappings.items():
                    csv_col_index = header_index.get(csv_column, -1)
                    raw_value = row[csv_col_index] if 0 <= csv_col_index < len(row) else ""
                    
                    # Apply transformation based on target field
                    if target_field == "VNK" and csv_column.lower() == "hiline_section":
                        value = self._extract_vnk(raw_value)
                    elif target_field == "NNK" and csv_column.lower() == "hiline_section":
                        value = self._extract_nnk(raw_value)
                    elif target_field == "KLASSE" and csv_column.lower() == "hiline_road":
                        value = self._extract_klasse(raw_value)
                    elif target_field == "NUMMER" and csv_column.lower() == "hiline_road":
                        value = self._extract_nummer(raw_value)
                    elif csv_column.lower() == "business_data" and target_field in [
                        "EFLI", "AFLI", "RISS", "ZWAUS", "ZWBIN", "ZWONA",
                        "ZWRSF", "ZWSCH", "ZWAFLI", "ZWBORD", "ZWEFLI",
                        "ZWRISS", "ZWWURZ", "GW", "GEB", "SUB"
                    ]:
                        value = self._extract_from_business_data(raw_value, target_field)
                    else:
                        # Default: use raw value (ID, FS, LAGE, etc.)
                        value = raw_value
                    
                    self.calculated_data[target_field].append(value)
            
            progress.update(total_rows, total_rows, "Complete!")
            progress.close()
            
            self.calc_status_label.config(text=f"✓ Calculated {len(active_mappings)} fields × {total_rows} rows")
            self.save_btn.config(state=tk.NORMAL)
            
            messagebox.showinfo(
                "Calculation Complete",
                f"Successfully calculated data:\n\n"
                f"• Fields: {len(active_mappings)}\n"
                f"• Rows: {total_rows}\n\n"
                f"You can now save to Access or CSV."
            )
            
        except Exception as e:
            progress.close()
            messagebox.showerror("Error", f"Calculation failed:\n{str(e)}")
            self.calculated_data = None
            self.calc_status_label.config(text="")
        
    def save_to_access(self):
        """Save to Access or CSV format."""
        if not self.calculated_data:
            messagebox.showwarning("Warning", "Please run Calculate first to process the data.")
            return
            
        format_choice = messagebox.askyesnocancel(
            "Choose Export Format",
            "Would you like to save as:\n\n"
            "YES = CSV (.csv) - Works everywhere\n"
            "NO = Access (.accdb) - Requires Access/ACE driver installed\n"
            "CANCEL = Cancel"
        )
        
        if format_choice is None:
            return
        
        if format_choice:
            file_path = filedialog.asksaveasfilename(
                title="Save CSV File",
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All files", "*.*")]
            )
            if not file_path:
                return
            self._save_as_csv(file_path)
        else:
            file_path = filedialog.asksaveasfilename(
                title="Save Access Database",
                defaultextension=".accdb",
                filetypes=[("Access Database", "*.accdb"), ("All files", "*.*")]
            )
            if not file_path:
                return
            self._save_as_access(file_path)
    
    def _save_as_csv(self, file_path):
        """Save calculated data as CSV file with semicolon delimiter for German compatibility."""
        try:
            fields_with_data = [f for f in self.calculated_data.keys() if self.calculated_data[f]]
            
            if not fields_with_data:
                messagebox.showwarning("Warning", "No data to export.")
                return
            
            num_rows = len(self.calculated_data[fields_with_data[0]])
            
            # Use semicolon as delimiter for German Excel compatibility
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                
                # Write header row
                writer.writerow(fields_with_data)
                
                # Write data rows
                for i in range(num_rows):
                    row = []
                    for field in fields_with_data:
                        value = self.calculated_data[field][i]
                        # Convert value to string, handle None
                        if value is None:
                            row.append("")
                        else:
                            row.append(str(value))
                    writer.writerow(row)
            
            messagebox.showinfo(
                "Success",
                f"CSV file saved to:\n{file_path}\n\n"
                f"Fields: {len(fields_with_data)}\n"
                f"Rows: {num_rows}\n"
                f"Delimiter: semicolon (;)"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV:\n{str(e)}")
    
    def _save_as_access(self, file_path):
        """Save calculated data as Access database."""
        try:
            import pyodbc
            
            fields_with_data = [f for f in self.calculated_data.keys() if self.calculated_data[f]]
            
            if not fields_with_data:
                messagebox.showwarning("Warning", "No data to export.")
                return
            
            num_rows = len(self.calculated_data[fields_with_data[0]])
            
            data_rows = []
            for i in range(num_rows):
                row = []
                for field in fields_with_data:
                    value = self.calculated_data[field][i]
                    if value is None:
                        row.append("")
                    else:
                        row.append(str(value))
                data_rows.append(row)
            
            if os.path.exists(file_path):
                os.remove(file_path)
            
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={file_path};'
            )
            
            conn = pyodbc.connect(conn_str, autocommit=True)
            cursor = conn.cursor()
            
            columns_sql = ', '.join([f'[{col}] TEXT(255)' for col in fields_with_data])
            cursor.execute(f'CREATE TABLE [MappedData] ({columns_sql})')
            
            if data_rows:
                placeholders = ', '.join(['?' for _ in fields_with_data])
                col_names = ', '.join([f'[{col}]' for col in fields_with_data])
                insert_sql = f'INSERT INTO [MappedData] ({col_names}) VALUES ({placeholders})'
                
                for row in data_rows:
                    cursor.execute(insert_sql, row)
            
            conn.commit()
            cursor.close()
            conn.close()
            
            messagebox.showinfo(
                "Success",
                f"Access database saved to:\n{file_path}\n\n"
                f"Table: MappedData\n"
                f"Fields: {len(fields_with_data)}\n"
                f"Rows: {len(data_rows)}"
            )
            
        except ImportError:
            messagebox.showerror("Error", "pyodbc is required.\n\nInstall with:\npip install pyodbc")
        except Exception as e:
            if "driver" in str(e).lower() or "data source" in str(e).lower():
                if messagebox.askyesno(
                    "Driver Not Found",
                    "Access database driver not found.\n\n"
                    "Would you like to save as CSV instead?"
                ):
                    csv_path = file_path.replace('.accdb', '.csv')
                    self._save_as_csv(csv_path)
            else:
                messagebox.showerror("Error", f"Failed to save:\n{str(e)}")


def main():
    root = tk.Tk()
    root.minsize(1000, 650)
    
    root.update_idletasks()
    width, height = 1100, 750
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    app = CSVFieldMapper(root)
    root.mainloop()
    
    return app.mappings


if __name__ == "__main__":
    main()
