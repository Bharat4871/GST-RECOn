import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re
from datetime import datetime
import os
import tempfile

class GSTReconciliationApp:
    """
    A Tkinter-based application for advanced GST reconciliation between GSTR-2A data
    and books of accounts purchase data.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced GST Reconciliation Tool")
        self.root.geometry("1300x850")
        self.root.state('zoomed') # Maximize the window on startup
        
        # Initialize settings with default values
        self.date_tolerance = 3  # Days
        self.amount_tolerance = 1.0  # Rupees
        self.auto_clean_gstin = True
        
        # Configure application style
        self.style = ttk.Style()
        self.style.configure("TNotebook.Tab", font=('Arial', 10, 'bold'), padding=[10, 5])
        self.style.configure("TButton", font=('Arial', 9))
        self.style.configure("Header.TLabel", font=('Arial', 12, 'bold'))
        self.style.configure("Bold.TLabel", font=('Arial', 10, 'bold'))
        
        # Create status bar at the bottom of the window
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize log text area early so log_message can be called from the start
        self.log_text = scrolledtext.ScrolledText(root, height=10, wrap=tk.WORD) 
        self.log_text.config(state=tk.DISABLED) # Make logs read-only initially
        # Note: self.log_text is created here, but will be packed into its frame in create_settings_tab

        # Try to set application icon (handle potential errors if icon file is missing)
        try:
            # It's good practice to provide a full path or ensure the icon is in the same directory
            self.root.iconbitmap("gst_icon.ico")
        except tk.TclError: # Catch specific Tkinter error for icon
            # Now self.log_text exists, so log_message can be called
            self.log_message("Warning: Application icon 'gst_icon.ico' not found.", error=True)
            pass # Icon not critical, so pass silently if it fails

        # Create main notebook (tabs container)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Initialize data containers as empty pandas DataFrames
        # Ensure they have the expected columns from the start to prevent KeyError later
        self.gstr2a_data = pd.DataFrame(columns=[
            'invoice_no', 'invoice_date', 'supplier_gstin', 'taxable_value', 
            'cgst', 'sgst', 'igst', 'total_amount', 'place_of_supply', 'match_key'
        ])
        self.books_data = pd.DataFrame(columns=[
            'invoice_no', 'invoice_date', 'supplier_gstin', 'taxable_value', 
            'cgst', 'sgst', 'igst', 'total_amount', 'place_of_supply', 
            'book_entry_date', 'match_key'
        ])
        self.reconciliation_results = None # To store reconciliation output
        
        # Create all application tabs
        self.create_gstr2a_tab()
        self.create_books_tab()
        self.create_gstr2a_manual_tab()
        self.create_books_manual_tab()
        self.create_reconciliation_tab()
        self.create_insights_tab()
        self.create_settings_tab() # This will now pack the already existing self.log_text
        
        # Initialize application logs
        self.log_message("Application started. Ready to load data.")
        
        # Store Excel template paths for future reference
        self.gstr2a_template_path = None
        self.books_template_path = None

    # --- TAB CREATION METHODS ---
    def create_gstr2a_tab(self):
        """Creates the GSTR-2A Import tab for loading data from files."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="GSTR-2A Import")
        
        header = ttk.Label(tab, text="GSTR-2A Data Import", style="Header.TLabel")
        header.pack(pady=15)
        
        # Frame for template download and open buttons
        template_frame = ttk.Frame(tab)
        template_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Button(template_frame, text="Download Template", 
                  command=self.download_gstr2a_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_frame, text="Open Template", 
                  command=self.open_gstr2a_template).pack(side=tk.LEFT, padx=5)
        
        # Frame for file selection
        file_frame = ttk.LabelFrame(tab, text="File Import")
        file_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(file_frame, text="Select File:").grid(row=0, column=0, padx=5, pady=10, sticky='w')
        self.gstr2a_file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.gstr2a_file_path, width=70).grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Button(file_frame, text="Browse", command=self.browse_gstr2a_file).grid(row=0, column=2, padx=5)
        ttk.Button(file_frame, text="Load Data", command=self.load_gstr2a_data).grid(row=0, column=3, padx=5)
        
        # Frame for column mapping (optional)
        map_frame = ttk.LabelFrame(tab, text="Column Mapping (Optional)")
        map_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Define standard column names and their corresponding UI labels
        mapping_fields = [
            ("Invoice No", "invoice_no"),
            ("Invoice Date", "invoice_date"),
            ("Supplier GSTIN", "supplier_gstin"),
            ("Taxable Value", "taxable_value"),
            ("CGST", "cgst"),
            ("SGST", "sgst"),
            ("IGST", "igst"),
            ("Total Amount", "total_amount"),
            ("Place of Supply", "place_of_supply")
        ]
        
        self.gstr2a_mapping_vars = {} # Dictionary to store mapping Tkinter variables
        for i, (label, key) in enumerate(mapping_fields):
            ttk.Label(map_frame, text=label + ":").grid(row=i, column=0, padx=5, pady=2, sticky='e')
            var = tk.StringVar()
            entry = ttk.Entry(map_frame, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky='w')
            self.gstr2a_mapping_vars[key] = var
        
        # Frame for data preview (Treeview)
        preview_frame = ttk.LabelFrame(tab, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Define columns for the Treeview
        columns = ("Invoice No", "Invoice Date", "Supplier GSTIN", "Taxable Value", 
                   "CGST", "SGST", "IGST", "Total Amount", "Place of Supply")
        self.gstr2a_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=12)
        
        # Add scrollbars to the Treeview
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.gstr2a_tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.gstr2a_tree.xview)
        self.gstr2a_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.gstr2a_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Configure Treeview column headings and widths
        for col in columns:
            self.gstr2a_tree.heading(col, text=col)
            self.gstr2a_tree.column(col, width=120, minwidth=80, anchor=tk.CENTER)
        
        # Configure grid weights for resizing
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Frame for data statistics
        stats_frame = ttk.Frame(tab)
        stats_frame.pack(fill=tk.X, padx=20, pady=5)
        
        self.gstr2a_stats = tk.StringVar(value="Records: 0 | Total Value: ₹0 | Total Tax: ₹0")
        ttk.Label(stats_frame, textvariable=self.gstr2a_stats, style="Bold.TLabel").pack(side=tk.LEFT)
        
        ttk.Button(tab, text="Clear Imported Data", command=self.clear_gstr2a_import).pack(pady=10)

    def create_books_tab(self):
        """Creates the Books Import tab for loading data from files."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Books Import")
        
        header = ttk.Label(tab, text="Books of Accounts Data Import", style="Header.TLabel")
        header.pack(pady=15)
        
        # Frame for template download and open buttons
        template_frame = ttk.Frame(tab)
        template_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Button(template_frame, text="Download Template", 
                  command=self.download_books_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_frame, text="Open Template", 
                  command=self.open_books_template).pack(side=tk.LEFT, padx=5)
        
        # Frame for file selection
        file_frame = ttk.LabelFrame(tab, text="File Import")
        file_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(file_frame, text="Select File:").grid(row=0, column=0, padx=5, pady=10, sticky='w')
        self.books_file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.books_file_path, width=70).grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Button(file_frame, text="Browse", command=self.browse_books_file).grid(row=0, column=2, padx=5)
        ttk.Button(file_frame, text="Load Data", command=self.load_books_data).grid(row=0, column=3, padx=5)
        
        # Frame for column mapping (optional)
        map_frame = ttk.LabelFrame(tab, text="Column Mapping (Optional)")
        map_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Define standard column names and their corresponding UI labels
        mapping_fields = [
            ("Invoice No", "invoice_no"),
            ("Invoice Date", "invoice_date"),
            ("Supplier GSTIN", "supplier_gstin"),
            ("Taxable Value", "taxable_value"),
            ("CGST", "cgst"),
            ("SGST", "sgst"),
            ("IGST", "igst"),
            ("Total Amount", "total_amount"),
            ("Place of Supply", "place_of_supply"),
            ("Book Entry Date", "book_entry_date")
        ]
        
        self.books_mapping_vars = {} # Dictionary to store mapping Tkinter variables
        for i, (label, key) in enumerate(mapping_fields):
            ttk.Label(map_frame, text=label + ":").grid(row=i, column=0, padx=5, pady=2, sticky='e')
            var = tk.StringVar()
            entry = ttk.Entry(map_frame, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky='w')
            self.books_mapping_vars[key] = var
        
        # Frame for data preview (Treeview)
        preview_frame = ttk.LabelFrame(tab, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Define columns for the Treeview
        columns = ("Invoice No", "Invoice Date", "Supplier GSTIN", "Taxable Value", 
                   "CGST", "SGST", "IGST", "Total Amount", "Place of Supply", "Book Entry Date")
        self.books_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=12)
        
        # Add scrollbars to the Treeview
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.books_tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.books_tree.xview)
        self.books_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.books_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Configure Treeview column headings and widths
        for col in columns:
            self.books_tree.heading(col, text=col)
            self.books_tree.column(col, width=120, minwidth=80, anchor=tk.CENTER)
        
        # Configure grid weights for resizing
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Frame for data statistics
        stats_frame = ttk.Frame(tab)
        stats_frame.pack(fill=tk.X, padx=20, pady=5)
        
        self.books_stats = tk.StringVar(value="Records: 0 | Total Value: ₹0 | Total Tax: ₹0")
        ttk.Label(stats_frame, textvariable=self.books_stats, style="Bold.TLabel").pack(side=tk.LEFT)
        
        ttk.Button(tab, text="Clear Imported Data", command=self.clear_books_import).pack(pady=10)

    def create_gstr2a_manual_tab(self):
        """Creates the GSTR-2A Manual Entry tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="GSTR-2A Manual")
        
        header = ttk.Label(tab, text="Manual GSTR-2A Data Entry", style="Header.TLabel")
        header.pack(pady=15)
        
        # Frame for data entry form
        form_frame = ttk.LabelFrame(tab, text="Add New Entry")
        form_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Define fields for manual entry with placeholders
        fields = [
            ("Invoice No", "e.g. INV-2023-001", "invoice_no"),
            ("Invoice Date (DD/MM/YYYY)", "e.g. 15/07/2023", "invoice_date"),
            ("Supplier GSTIN", "e.g. 22AAAAA0000A1Z5", "supplier_gstin"),
            ("Taxable Value", "e.g. 10000.00", "taxable_value"),
            ("CGST", "e.g. 900.00", "cgst"),
            ("SGST", "e.g. 900.00", "sgst"),
            ("IGST", "e.g. 0.00", "igst"),
            ("Total Amount", "e.g. 11800.00", "total_amount"),
            ("Place of Supply", "e.g. 07", "place_of_supply")
        ]
        
        self.gstr2a_entries = {} # Dictionary to store entry widgets
        for i, (label, placeholder, key) in enumerate(fields):
            frame = ttk.Frame(form_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            ttk.Label(frame, text=label, width=25, anchor='e').pack(side=tk.LEFT, padx=5)
            entry = ttk.Entry(frame)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            entry.insert(0, placeholder) # Set placeholder text
            self.gstr2a_entries[key] = entry
        
        # Buttons for adding and clearing entries
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Add Entry", command=self.add_gstr2a_manual).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Clear Form", command=self.clear_gstr2a_form).pack(side=tk.LEFT, padx=10)
        
        # Frame for data preview (Treeview) of manual entries
        preview_frame = ttk.LabelFrame(tab, text="Manual Entries")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        columns = ("Invoice No", "Invoice Date", "Supplier GSTIN", "Taxable Value", 
                   "CGST", "SGST", "IGST", "Total Amount", "Place of Supply")
        self.gstr2a_manual_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=12)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.gstr2a_manual_tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.gstr2a_manual_tree.xview)
        self.gstr2a_manual_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.gstr2a_manual_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        for col in columns:
            self.gstr2a_manual_tree.heading(col, text=col)
            self.gstr2a_manual_tree.column(col, width=120, minwidth=80, anchor=tk.CENTER)
        
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Action buttons for manual entries
        action_frame = ttk.Frame(tab)
        action_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Button(action_frame, text="Delete Selected", command=self.delete_gstr2a_manual).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Clear All Entries", command=self.clear_gstr2a_manual).pack(side=tk.LEFT, padx=5)

    def create_books_manual_tab(self):
        """Creates the Books Manual Entry tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Books Manual")
        
        header = ttk.Label(tab, text="Manual Books of Accounts Data Entry", style="Header.TLabel")
        header.pack(pady=15)
        
        # Frame for data entry form
        form_frame = ttk.LabelFrame(tab, text="Add New Entry")
        form_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Define fields for manual entry with placeholders
        fields = [
            ("Invoice No", "e.g. INV-2023-001", "invoice_no"),
            ("Invoice Date (DD/MM/YYYY)", "e.g. 15/07/2023", "invoice_date"),
            ("Supplier GSTIN", "e.g. 22AAAAA0000A1Z5", "supplier_gstin"),
            ("Taxable Value", "e.g. 10000.00", "taxable_value"),
            ("CGST", "e.g. 900.00", "cgst"),
            ("SGST", "e.g. 900.00", "sgst"),
            ("IGST", "e.g. 0.00", "igst"),
            ("Total Amount", "e.g. 11800.00", "total_amount"),
            ("Place of Supply", "e.g. 07", "place_of_supply"),
            ("Book Entry Date (DD/MM/YYYY)", "e.g. 18/07/2023", "book_entry_date")
        ]
        
        self.books_entries = {} # Dictionary to store entry widgets
        for i, (label, placeholder, key) in enumerate(fields):
            frame = ttk.Frame(form_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            ttk.Label(frame, text=label, width=25, anchor='e').pack(side=tk.LEFT, padx=5)
            entry = ttk.Entry(frame)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            entry.insert(0, placeholder) # Set placeholder text
            self.books_entries[key] = entry
        
        # Buttons for adding and clearing entries
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Add Entry", command=self.add_books_manual).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Clear Form", command=self.clear_books_form).pack(side=tk.LEFT, padx=10)
        
        # Frame for data preview (Treeview) of manual entries
        preview_frame = ttk.LabelFrame(tab, text="Manual Entries")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        columns = ("Invoice No", "Invoice Date", "Supplier GSTIN", "Taxable Value", 
                   "CGST", "SGST", "IGST", "Total Amount", "Place of Supply", "Book Entry Date")
        self.books_manual_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=12)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.books_manual_tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.books_manual_tree.xview)
        self.books_manual_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.books_manual_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        for col in columns:
            self.books_manual_tree.heading(col, text=col)
            self.books_manual_tree.column(col, width=120, minwidth=80, anchor=tk.CENTER)
        
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Action buttons for manual entries
        action_frame = ttk.Frame(tab)
        action_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Button(action_frame, text="Delete Selected", command=self.delete_books_manual).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Clear All Entries", command=self.clear_books_manual).pack(side=tk.LEFT, padx=5)

    def create_reconciliation_tab(self):
        """Creates the Reconciliation tab to display results."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Reconciliation")
        
        # Control panel for reconciliation actions
        control_frame = ttk.LabelFrame(tab, text="Reconciliation Control")
        control_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(control_frame, text="Status:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.recon_status = tk.StringVar(value="Ready")
        ttk.Label(control_frame, textvariable=self.recon_status, foreground="blue", 
                 font=('Arial', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        ttk.Button(control_frame, text="Run Reconciliation", command=self.run_reconciliation).grid(row=0, column=2, padx=10)
        ttk.Button(control_frame, text="Export Results", command=self.export_results).grid(row=0, column=3, padx=10)
        ttk.Button(control_frame, text="Show Data Summary", command=self.show_data_summary).grid(row=0, column=4, padx=10)
        
        # Results display area
        results_frame = ttk.LabelFrame(tab, text="Reconciliation Results")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Notebook for different discrepancy types
        self.results_notebook = ttk.Notebook(results_frame)
        self.results_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Define discrepancy types and their display names
        discrepancy_types = {
            "all": "All Discrepancies",
            "missing_in_books": "Missing in Books",
            "missing_in_gstr2a": "Missing in GSTR-2A",
            "date_mismatch": "Date Mismatch",
            "amount_mismatch": "Amount/Tax Mismatch",
            "gstin_mismatch": "GSTIN Mismatch",
            "duplicates": "Duplicates"
        }
        
        self.discrepancy_trees = {} # Dictionary to store Treeview widgets for each discrepancy type
        for key, name in discrepancy_types.items():
            frame = ttk.Frame(self.results_notebook)
            self.results_notebook.add(frame, text=name)
            
            # Create Treeview for each discrepancy type
            columns = ("Invoice No", "Source", "Issue Type", "GSTR-2A Date", "Books Date", 
                       "GSTR-2A GSTIN", "Books GSTIN", "Amount Diff", "Tax Diff", "Details")
            tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)
            
            # Add scrollbars
            vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            
            tree.grid(row=0, column=0, sticky='nsew')
            vsb.grid(row=0, column=1, sticky='ns')
            hsb.grid(row=1, column=0, sticky='ew')
            
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=120, minwidth=80, anchor=tk.CENTER)
            
            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)
            
            self.discrepancy_trees[key] = tree

    def create_insights_tab(self):
        """Creates the Insights tab for displaying charts and a summary report."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Insights")
        
        header = ttk.Label(tab, text="Reconciliation Insights", style="Header.TLabel")
        header.pack(pady=15)
        
        # Frame for matplotlib charts
        charts_frame = ttk.Frame(tab)
        charts_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Create a figure and subplots for charts
        self.fig, self.axs = plt.subplots(2, 2, figsize=(12, 10))
        self.fig.tight_layout(pad=5.0) # Adjust layout to prevent overlap
        self.canvas = FigureCanvasTkAgg(self.fig, master=charts_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Frame for summary report (scrolled text)
        summary_frame = ttk.LabelFrame(tab, text="Summary Report")
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.summary_text = scrolledtext.ScrolledText(summary_frame, height=12, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.summary_text.config(state=tk.DISABLED) # Make text read-only

    def create_settings_tab(self):
        """Creates the Settings tab for configuring reconciliation parameters and managing data."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settings")
        
        header = ttk.Label(tab, text="Reconciliation Settings", style="Header.TLabel")
        header.pack(pady=15)
        
        # Frame for tolerance settings
        settings_frame = ttk.LabelFrame(tab, text="Tolerance Settings")
        settings_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Date tolerance input
        ttk.Label(settings_frame, text="Date Difference Tolerance (days):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.date_tol_var = tk.StringVar(value=str(self.date_tolerance))
        ttk.Entry(settings_frame, textvariable=self.date_tol_var, width=10).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Amount tolerance input
        ttk.Label(settings_frame, text="Amount/Tax Difference Tolerance (₹):").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.amount_tol_var = tk.StringVar(value=str(self.amount_tolerance))
        ttk.Entry(settings_frame, textvariable=self.amount_tol_var, width=10).grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Auto-clean GSTIN checkbox
        self.auto_clean_var = tk.IntVar(value=1 if self.auto_clean_gstin else 0)
        ttk.Checkbutton(settings_frame, text="Automatically clean GSTIN numbers", 
                       variable=self.auto_clean_var).grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='w')
        
        # Save settings button
        ttk.Button(settings_frame, text="Save Settings", command=self.save_settings).grid(row=3, column=0, padx=5, pady=10, sticky='w')
        
        # Data management frame
        data_frame = ttk.LabelFrame(tab, text="Data Management")
        data_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Button(data_frame, text="Clear All GSTR-2A Data", command=self.clear_all_gstr2a).pack(side=tk.LEFT, padx=10, pady=5)
        ttk.Button(data_frame, text="Clear All Books Data", command=self.clear_all_books).pack(side=tk.LEFT, padx=10, pady=5)
        ttk.Button(data_frame, text="Export All Data", command=self.export_all_data).pack(side=tk.LEFT, padx=10, pady=5)
        
        # Application logs display
        log_frame = ttk.LabelFrame(tab, text="Application Logs")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # self.log_text is already created in __init__, just pack it now
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # self.log_text.config(state=tk.DISABLED) # This was already set in __init__
        
        self.log_message("Settings tab initialized")

    # --- EXCEL TEMPLATE METHODS ---
    def download_gstr2a_template(self):
        """Generates and downloads an Excel template for GSTR-2A data."""
        columns = [
            'Invoice No', 'Invoice Date', 'Supplier GSTIN', 
            'Taxable Value', 'CGST', 'SGST', 'IGST', 
            'Total Amount', 'Place of Supply'
        ]
        df = pd.DataFrame(columns=columns)
        
        # Create a temporary file path for the template
        # Using tempfile.mkstemp to get a unique file name and path securely
        fd, self.gstr2a_template_path = tempfile.mkstemp(suffix='_GSTR2A_Template.xlsx')
        os.close(fd) # Close the file descriptor immediately as we will use pandas to write
        
        # Save the empty DataFrame to Excel
        df.to_excel(self.gstr2a_template_path, index=False)
        
        # Add sample data to the template
        sample_data = [
            ['INV-2023-001', '15/07/2023', '22AAAAA0000A1Z5', 10000.00, 900.00, 900.00, 0.00, 11800.00, '07'],
            ['INV-2023-002', '18/07/2023', '33BBBBB0000B2Z6', 15000.00, 1350.00, 1350.00, 0.00, 17700.00, '07']
        ]
        sample_df = pd.DataFrame(sample_data, columns=columns)
        # Use openpyxl engine to append data to an existing sheet
        try:
            with pd.ExcelWriter(self.gstr2a_template_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                sample_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, header=False)
        except Exception as e:
            self.log_message(f"Error adding sample data to GSTR-2A template: {e}", error=True)

        self.log_message(f"GSTR-2A template created at: {self.gstr2a_template_path}")
        messagebox.showinfo("Template Created", 
                          "GSTR-2A template downloaded successfully. Please fill in your data and import it.")
    
    def open_gstr2a_template(self):
        """Opens the downloaded GSTR-2A template."""
        if not self.gstr2a_template_path or not os.path.exists(self.gstr2a_template_path):
            self.log_message("GSTR-2A template not found, attempting to download.", error=False)
            self.download_gstr2a_template() # Download if not already present
        
        if self.gstr2a_template_path and os.path.exists(self.gstr2a_template_path):
            try:
                os.startfile(self.gstr2a_template_path) # Open the file using default application
            except Exception as e:
                messagebox.showerror("Error", f"Could not open the template file: {e}. Please try again.")
                self.log_message(f"Error opening GSTR-2A template: {e}", error=True)
        else:
            messagebox.showerror("Error", "GSTR-2A template could not be created or found.")
            self.log_message("Failed to open GSTR-2A template: File not available.", error=True)

    def download_books_template(self):
        """Generates and downloads an Excel template for Books data."""
        columns = [
            'Invoice No', 'Invoice Date', 'Supplier GSTIN', 
            'Taxable Value', 'CGST', 'SGST', 'IGST', 
            'Total Amount', 'Place of Supply', 'Book Entry Date'
        ]
        df = pd.DataFrame(columns=columns)
        
        # Create a temporary file path for the template
        fd, self.books_template_path = tempfile.mkstemp(suffix='_Books_Template.xlsx')
        os.close(fd) # Close the file descriptor immediately
        
        # Save the empty DataFrame to Excel
        df.to_excel(self.books_template_path, index=False)
        
        # Add sample data to the template
        sample_data = [
            ['INV-2023-001', '15/07/2023', '22AAAAA0000A1Z5', 10000.00, 900.00, 900.00, 0.00, 11800.00, '07', '17/07/2023'],
            ['INV-2023-003', '20/07/2023', '44CCCCC0000C3Z7', 20000.00, 1800.00, 1800.00, 0.00, 23600.00, '07', '22/07/2023']
        ]
        sample_df = pd.DataFrame(sample_data, columns=columns)
        # Use openpyxl engine to append data to an existing sheet
        try:
            with pd.ExcelWriter(self.books_template_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                sample_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, header=False)
        except Exception as e:
            self.log_message(f"Error adding sample data to Books template: {e}", error=True)

        self.log_message(f"Books template created at: {self.books_template_path}")
        messagebox.showinfo("Template Created", 
                          "Books template downloaded successfully. Please fill in your data and import it.")
    
    def open_books_template(self):
        """Opens the downloaded Books template."""
        if not self.books_template_path or not os.path.exists(self.books_template_path):
            self.log_message("Books template not found, attempting to download.", error=False)
            self.download_books_template() # Download if not already present
        
        if self.books_template_path and os.path.exists(self.books_template_path):
            try:
                os.startfile(self.books_template_path) # Open the file using default application
            except Exception as e:
                messagebox.showerror("Error", f"Could not open the template file: {e}. Please try again.")
                self.log_message(f"Error opening Books template: {e}", error=True)
        else:
            messagebox.showerror("Error", "Books template could not be created or found.")
            self.log_message("Failed to open Books template: File not available.", error=True)
    
    # --- DATA IMPORT METHODS ---
    def browse_gstr2a_file(self):
        """Opens a file dialog to select the GSTR-2A data file."""
        file_path = filedialog.askopenfilename(
            title="Select GSTR-2A File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.gstr2a_file_path.set(file_path)
            self.log_message(f"Selected GSTR-2A file: {file_path}")

    def browse_books_file(self):
        """Opens a file dialog to select the Books of Accounts data file."""
        file_path = filedialog.askopenfilename(
            title="Select Books of Accounts File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.books_file_path.set(file_path)
            self.log_message(f"Selected Books file: {file_path}")

    def load_gstr2a_data(self):
        """Loads and processes GSTR-2A data from the selected file."""
        file_path = self.gstr2a_file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file first.")
            self.log_message("No GSTR-2A file selected or file does not exist.", error=True)
            return
        
        try:
            # Read file based on extension
            if file_path.lower().endswith(('.csv')):
                df = pd.read_csv(file_path)
            elif file_path.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select an Excel or CSV file.")
                self.log_message("Unsupported GSTR-2A file format.", error=True)
                return

            self.log_message(f"Loaded GSTR-2A file with {len(df)} records.")
            
            # Get custom mapping from UI entries
            custom_mapping = {key: var.get().strip() for key, var in self.gstr2a_mapping_vars.items() if var.get().strip()}
            
            # Clean and transform the loaded data
            df = self.clean_and_transform_data(df, "GSTR-2A", custom_mapping)
            
            # Concatenate with existing GSTR-2A data
            # Use pd.concat with ignore_index=True to handle cases where initial DFs are empty
            self.gstr2a_data = pd.concat([self.gstr2a_data, df], ignore_index=True)
            
            # Update UI elements
            self.update_treeview(self.gstr2a_tree, self.gstr2a_data) # Update main GSTR-2A preview
            self.update_treeview(self.gstr2a_manual_tree, self.gstr2a_data) # Also update manual tree as it shows combined data
            self.update_gstr2a_stats()
            
            self.log_message(f"Successfully processed {len(df)} GSTR-2A records.")
            messagebox.showinfo("Success", "GSTR-2A data loaded successfully.")
            
        except Exception as e:
            self.log_message(f"Error loading GSTR-2A data: {str(e)}", error=True)
            messagebox.showerror("Error", f"Failed to load GSTR-2A data: {str(e)}")

    def load_books_data(self):
        """Loads and processes Books of Accounts data from the selected file."""
        file_path = self.books_file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file first.")
            self.log_message("No Books file selected or file does not exist.", error=True)
            return
        
        try:
            # Read file based on extension
            if file_path.lower().endswith(('.csv')):
                df = pd.read_csv(file_path)
            elif file_path.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select an Excel or CSV file.")
                self.log_message("Unsupported Books file format.", error=True)
                return
            
            self.log_message(f"Loaded Books file with {len(df)} records.")
            
            # Get custom mapping from UI entries
            custom_mapping = {key: var.get().strip() for key, var in self.books_mapping_vars.items() if var.get().strip()}
            
            # Clean and transform the loaded data
            df = self.clean_and_transform_data(df, "Books", custom_mapping)
            
            # Concatenate with existing Books data
            self.books_data = pd.concat([self.books_data, df], ignore_index=True)
            
            # Update UI elements
            self.update_treeview(self.books_tree, self.books_data) # Update main Books preview
            self.update_treeview(self.books_manual_tree, self.books_data) # Also update manual tree
            self.update_books_stats()
            
            self.log_message(f"Successfully processed {len(df)} Books records.")
            messagebox.showinfo("Success", "Books data loaded successfully.")
            
        except Exception as e:
            self.log_message(f"Error loading Books data: {str(e)}", error=True)
            messagebox.showerror("Error", f"Failed to load Books data: {str(e)}")

    def clean_and_transform_data(self, df, source, custom_mapping=None):
        """
        Cleans and transforms the input DataFrame to a standardized format.
        Applies custom column mapping, standardizes column names, converts data types,
        cleans GSTINs, and creates a match key.
        """
        if df.empty:
            # Return a DataFrame with standard columns to maintain schema consistency
            standard_columns = [
                'invoice_no', 'invoice_date', 'supplier_gstin', 'taxable_value', 
                'cgst', 'sgst', 'igst', 'total_amount', 'place_of_supply'
            ]
            if source == "Books":
                standard_columns.append('book_entry_date')
            standard_columns.append('match_key')
            return pd.DataFrame(columns=standard_columns)

        # Create a mapping of expected standard column names
        # Keys are standardized internal names, values are common variations
        standard_columns_map = {
            'invoice_no': ['invoice no', 'invoiceno', 'bill no', 'billno', 'invoice'],
            'invoice_date': ['invoice date', 'invoicedate', 'bill date', 'billdate', 'date'],
            'supplier_gstin': ['supplier gstin', 'suppliergstin', 'gstin', 'party gstin', 'receiver gstin'],
            'taxable_value': ['taxable value', 'taxablevalue', 'value', 'net amount'],
            'cgst': ['cgst', 'central tax'],
            'sgst': ['sgst', 'state tax'],
            'igst': ['igst', 'integrated tax'],
            'total_amount': ['total amount', 'totalamount', 'amount', 'gross amount'],
            'place_of_supply': ['place of supply', 'placeofsupply', 'pos']
        }
        
        if source == "Books":
            standard_columns_map['book_entry_date'] = ['book entry date', 'bookentrydate', 'entry date', 'accounting date']

        # Normalize existing DataFrame columns for easier matching
        original_cols = df.columns.tolist()
        df.columns = df.columns.str.strip().str.lower().str.replace(r'[^a-z0-9]', '', regex=True)
        normalized_df_cols = df.columns.tolist()

        # Reverse the standard_columns_map for easy lookup from actual column names to standard names
        reverse_standard_map = {}
        for std_name, variations in standard_columns_map.items():
            for var in variations:
                reverse_standard_map[var.replace(' ', '').lower()] = std_name

        # Apply custom mapping first
        if custom_mapping:
            # Normalize custom mapping keys and values
            normalized_custom_mapping = {}
            for std_key, user_col_display in custom_mapping.items():
                normalized_user_col = user_col_display.strip().lower().replace(' ', '').replace(r'[^a-z0-9]', '', regex=True)
                # Find the actual column name in the DataFrame's normalized columns
                if normalized_user_col in normalized_df_cols:
                    # Get the original column name corresponding to the normalized one
                    original_col_idx = normalized_df_cols.index(normalized_user_col)
                    original_col_name = original_cols[original_col_idx]
                    normalized_custom_mapping[original_col_name] = std_key
                else:
                    self.log_message(f"Warning: Custom mapped column '{user_col_display}' for '{std_key}' not found in data. Skipping.", error=True)

            # Rename columns based on custom mapping
            df = df.rename(columns=normalized_custom_mapping)
            # Re-normalize columns after custom renaming to ensure consistency for standard mapping
            df.columns = df.columns.str.strip().str.lower().str.replace(r'[^a-z0-9]', '', regex=True)


        # Now, map remaining columns to standard names
        # Create a new dictionary for renaming to avoid modifying df.columns directly during iteration
        rename_dict = {}
        for col in df.columns:
            normalized_col_for_lookup = col.replace(' ', '').lower()
            if normalized_col_for_lookup in reverse_standard_map:
                rename_dict[col] = reverse_standard_map[normalized_col_for_lookup]
            # else: keep original if no standard match, but for reconciliation, we only care about standard ones
        df = df.rename(columns=rename_dict)

        # Ensure all standard columns exist, add with None if missing
        for std_name in standard_columns_map.keys():
            if std_name not in df.columns:
                df[std_name] = None
        
        # Convert date columns
        date_cols = ['invoice_date']
        if source == "Books":
            date_cols.append('book_entry_date')
            
        for col in date_cols:
            if col in df.columns:
                # Convert to datetime, coercing errors to NaT (Not a Time)
                # Try dayfirst=True for DD/MM/YYYY, then mixed for other common formats
                df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
                # If still NaT for some, try with mixed format
                df.loc[df[col].isna(), col] = pd.to_datetime(df.loc[df[col].isna(), col], errors='coerce', format='mixed')
            else:
                df[col] = pd.NaT # Ensure column exists and is NaT if no data

        # Convert numeric columns
        numeric_cols = ['taxable_value', 'cgst', 'sgst', 'igst', 'total_amount']
        for col in numeric_cols:
            if col in df.columns:
                # Convert to string first to handle mixed types, then remove non-numeric chars
                # and convert to numeric, filling NaNs with 0.0
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(r'[^\d.]', '', regex=True),
                    errors='coerce'
                ).fillna(0.0).astype(float)
            else:
                df[col] = 0.0 # Ensure column exists and is 0.0 if no data

        # Clean GSTIN
        if 'supplier_gstin' in df.columns and self.auto_clean_gstin:
            df['supplier_gstin'] = df['supplier_gstin'].apply(self.clean_gstin)
        else:
            # Ensure 'supplier_gstin' column exists and is string type, fill NaN with empty string
            df['supplier_gstin'] = df.get('supplier_gstin', pd.Series(dtype=str)).fillna('').astype(str)

        # Create 'match_key' for reconciliation
        # Ensure 'invoice_no' and 'supplier_gstin' are strings before concatenation
        df['invoice_no'] = df.get('invoice_no', pd.Series(dtype=str)).astype(str)
        df['supplier_gstin'] = df.get('supplier_gstin', pd.Series(dtype=str)).astype(str)
        
        # Create match key. Handle cases where either might be missing or empty.
        df['match_key'] = df['invoice_no'] + '_' + df['supplier_gstin']
        # If invoice_no or supplier_gstin is empty/None/NaN, set match_key to NaN
        df.loc[
            (df['invoice_no'].isin(['', 'None', 'nan', None])) | 
            (df['supplier_gstin'].isin(['', 'None', 'nan', None])), 
            'match_key'
        ] = np.nan

        # Fill missing place of supply
        if 'place_of_supply' in df.columns:
            df['place_of_supply'] = df['place_of_supply'].fillna('').astype(str)
        else:
            df['place_of_supply'] = '' # Ensure column exists

        # Select and reorder only the standard columns for the output DataFrame
        # This ensures consistent schema for reconciliation
        final_cols_order = list(standard_columns_map.keys())
        if 'match_key' not in final_cols_order:
            final_cols_order.append('match_key')
        
        # Filter df to only include columns that are in final_cols_order and actually exist in df
        # Reindex to ensure the order
        df = df[[col for col in final_cols_order if col in df.columns]].reindex(columns=final_cols_order)
        
        return df

    def clean_gstin(self, gstin):
        """
        Cleans and standardizes a GSTIN string.
        Removes non-alphanumeric characters, truncates/pads to 15 characters,
        and converts to uppercase.
        """
        if pd.isna(gstin) or str(gstin).strip().lower() in ['', 'nan', 'none']:
            return "" # Return empty string for NaN or empty inputs
        
        try:
            gstin = str(gstin).strip().upper()
            # Remove any character that is not an uppercase letter or a digit
            gstin = re.sub(r'[^A-Z0-9]', '', gstin)
            
            # GSTINs are 15 characters long. Handle common issues.
            if len(gstin) > 15:
                gstin = gstin[:15]  # Truncate if longer than 15
            elif len(gstin) < 15:
                # Pad with 'X' to indicate potential incompleteness for GSTINs shorter than 15.
                gstin = gstin.ljust(15, 'X') 
            
            # Basic validation: check if it's exactly 15 chars and starts with 2 digits (state code)
            # This is a basic check, full GSTIN validation is complex
            if len(gstin) == 15 and gstin[:2].isdigit() and re.fullmatch(r'[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}', gstin):
                 return gstin
            
            # If basic validation fails, return the cleaned string with a warning or a specific invalid marker
            self.log_message(f"Warning: GSTIN '{gstin}' appears to be invalid after cleaning.", error=False)
            return gstin 
        except Exception as e:
            self.log_message(f"Error cleaning GSTIN '{gstin}': {e}", error=True)
            return "" # Return empty string on error, safer than "INVALID_GSTIN" for matching

    # --- MANUAL ENTRY METHODS ---
    def add_gstr2a_manual(self):
        """Adds a new manual GSTR-2A entry to the data."""
        try:
            entry_data = {}
            # Define placeholders to ignore
            placeholders = {
                "invoice_no": "e.g. INV-2023-001",
                "invoice_date": "e.g. 15/07/2023",
                "supplier_gstin": "e.g. 22AAAAA0000A1Z5", 
                "taxable_value": "e.g. 10000.00", 
                "cgst": "e.g. 900.00", 
                "sgst": "e.g. 900.00", 
                "igst": "e.g. 0.00", 
                "total_amount": "e.g. 11800.00", 
                "place_of_supply": "e.g. 07"
            }

            # Collect data from each entry widget
            for key, entry in self.gstr2a_entries.items():
                value = entry.get().strip()
                # Only add if value is not a placeholder
                if value != placeholders.get(key, ""):
                    entry_data[key] = value
                else:
                    entry_data[key] = "" # Ensure empty string for placeholders

            # Create a DataFrame from the single entry
            df = pd.DataFrame([entry_data])
            
            # Clean and transform the new entry
            df = self.clean_and_transform_data(df, "GSTR-2A")
            
            # Check if the cleaned DataFrame is empty (e.g., if all critical fields were missing)
            if df.empty or df.iloc[0].isnull().all(): # Check if the first row is all nulls
                messagebox.showerror("Error", "No valid data to add. Please fill in required fields.")
                self.log_message("Attempted to add empty GSTR-2A manual entry.", error=True)
                return

            # Concatenate with existing GSTR-2A data
            self.gstr2a_data = pd.concat([self.gstr2a_data, df], ignore_index=True)
            
            # Update UI elements for manual entries and stats
            self.update_treeview(self.gstr2a_manual_tree, self.gstr2a_data)
            self.update_treeview(self.gstr2a_tree, self.gstr2a_data) # Update main import tree too
            self.update_gstr2a_stats()
            
            self.clear_gstr2a_form() # Clear the form after successful addition
            
            self.log_message("Added manual GSTR-2A entry.")
            messagebox.showinfo("Success", "Entry added to GSTR-2A data.")
            
        except Exception as e:
            self.log_message(f"Error adding manual GSTR-2A entry: {str(e)}", error=True)
            messagebox.showerror("Error", f"Failed to add entry: {str(e)}")

    def add_books_manual(self):
        """Adds a new manual Books entry to the data."""
        try:
            entry_data = {}
            # Define placeholders to ignore
            placeholders = {
                "invoice_no": "e.g. INV-2023-001",
                "invoice_date": "e.g. 15/07/2023",
                "supplier_gstin": "e.g. 22AAAAA0000A1Z5", 
                "taxable_value": "e.g. 10000.00", 
                "cgst": "e.g. 900.00", 
                "sgst": "e.g. 900.00", 
                "igst": "e.g. 0.00", 
                "total_amount": "e.g. 11800.00", 
                "place_of_supply": "e.g. 07", 
                "book_entry_date": "e.g. 18/07/2023"
            }

            # Collect data from each entry widget
            for key, entry in self.books_entries.items():
                value = entry.get().strip()
                # Only add if value is not a placeholder
                if value != placeholders.get(key, ""):
                    entry_data[key] = value
                else:
                    entry_data[key] = "" # Ensure empty string for placeholders
            
            # Create a DataFrame from the single entry
            df = pd.DataFrame([entry_data])
            
            # Clean and transform the new entry
            df = self.clean_and_transform_data(df, "Books")

            # Check if the cleaned DataFrame is empty
            if df.empty or df.iloc[0].isnull().all(): # Check if the first row is all nulls
                messagebox.showerror("Error", "No valid data to add. Please fill in required fields.")
                self.log_message("Attempted to add empty Books manual entry.", error=True)
                return
            
            # Concatenate with existing Books data
            self.books_data = pd.concat([self.books_data, df], ignore_index=True)
            
            # Update UI elements for manual entries and stats
            self.update_treeview(self.books_manual_tree, self.books_data)
            self.update_treeview(self.books_tree, self.books_data) # Update main import tree too
            self.update_books_stats()
            
            self.clear_books_form() # Clear the form after successful addition
            
            self.log_message("Added manual Books entry.")
            messagebox.showinfo("Success", "Entry added to Books data.")
            
        except Exception as e:
            self.log_message(f"Error adding manual Books entry: {str(e)}", error=True)
            messagebox.showerror("Error", f"Failed to add entry: {str(e)}")

    def clear_gstr2a_form(self):
        """Clears the GSTR-2A manual entry form and restores placeholders."""
        fields = {
            "invoice_no": "e.g. INV-2023-001",
            "invoice_date": "e.g. 15/07/2023",
            "supplier_gstin": "e.g. 22AAAAA0000A1Z5",
            "taxable_value": "e.g. 10000.00",
            "cgst": "e.g. 900.00",
            "sgst": "e.g. 900.00",
            "igst": "e.g. 0.00",
            "total_amount": "e.g. 11800.00",
            "place_of_supply": "e.g. 07"
        }
        for key, entry in self.gstr2a_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, fields.get(key, "")) # Restore placeholder

    def clear_books_form(self):
        """Clears the Books manual entry form and restores placeholders."""
        fields = {
            "invoice_no": "e.g. INV-2023-001",
            "invoice_date": "e.g. 15/07/2023",
            "supplier_gstin": "e.g. 22AAAAA0000A1Z5",
            "taxable_value": "e.g. 10000.00",
            "cgst": "e.g. 900.00",
            "sgst": "e.g. 900.00",
            "igst": "e.g. 0.00",
            "total_amount": "e.g. 11800.00",
            "place_of_supply": "e.g. 07",
            "book_entry_date": "e.g. 18/07/2023"
        }
        for key, entry in self.books_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, fields.get(key, "")) # Restore placeholder

    def delete_gstr2a_manual(self):
        """Deletes selected manual GSTR-2A entries from the data."""
        selected_items = self.gstr2a_manual_tree.selection()
        if not selected_items:
            messagebox.showinfo("Info", "Please select one or more entries to delete.")
            return
            
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {len(selected_items)} selected entries?"):
            try:
                # Collect match_keys of selected items for robust deletion
                match_keys_to_delete = []
                for item in selected_items:
                    # Assuming 'Invoice No' and 'Supplier GSTIN' are sufficient to form the match_key
                    # and are present in the Treeview's values in the correct order.
                    # The `clean_and_transform_data` creates `match_key` as `invoice_no`_`supplier_gstin`.
                    invoice_no = self.gstr2a_manual_tree.item(item, 'values')[0]
                    supplier_gstin = self.gstr2a_manual_tree.item(item, 'values')[2] # Assuming GSTIN is the 3rd column (index 2)
                    
                    # Re-clean GSTIN to match how it's stored in the DataFrame's match_key
                    cleaned_gstin = self.clean_gstin(supplier_gstin.replace('₹', '').strip()) # Remove currency symbol if present
                    
                    match_key = str(invoice_no).strip() + '_' + cleaned_gstin
                    match_keys_to_delete.append(match_key)
                
                prev_count = len(self.gstr2a_data)
                # Filter out rows where 'match_key' is in the list of match_keys_to_delete
                self.gstr2a_data = self.gstr2a_data[~self.gstr2a_data['match_key'].isin(match_keys_to_delete)].reset_index(drop=True)
                deleted_count = prev_count - len(self.gstr2a_data)
                
                if deleted_count > 0:
                    self.update_treeview(self.gstr2a_manual_tree, self.gstr2a_data)
                    self.update_treeview(self.gstr2a_tree, self.gstr2a_data) # Update main import tree
                    self.update_gstr2a_stats()
                    self.log_message(f"Deleted {deleted_count} GSTR-2A entries.")
                    messagebox.showinfo("Success", f"{deleted_count} entries deleted successfully.")
                else:
                    messagebox.showinfo("Info", "No matching entries found to delete.")
                
            except Exception as e:
                self.log_message(f"Error deleting GSTR-2A entries: {str(e)}", error=True)
                messagebox.showerror("Error", f"Failed to delete entries: {str(e)}")

    def delete_books_manual(self):
        """Deletes selected manual Books entries from the data."""
        selected_items = self.books_manual_tree.selection()
        if not selected_items:
            messagebox.showinfo("Info", "Please select one or more entries to delete.")
            return
            
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {len(selected_items)} selected entries?"):
            try:
                match_keys_to_delete = []
                for item in selected_items:
                    invoice_no = self.books_manual_tree.item(item, 'values')[0]
                    supplier_gstin = self.books_manual_tree.item(item, 'values')[2] # Assuming GSTIN is the 3rd column (index 2)
                    
                    # Re-clean GSTIN to match how it's stored in the DataFrame's match_key
                    cleaned_gstin = self.clean_gstin(supplier_gstin.replace('₹', '').strip())
                    
                    match_key = str(invoice_no).strip() + '_' + cleaned_gstin
                    match_keys_to_delete.append(match_key)
                
                prev_count = len(self.books_data)
                self.books_data = self.books_data[~self.books_data['match_key'].isin(match_keys_to_delete)].reset_index(drop=True)
                deleted_count = prev_count - len(self.books_data)
                
                if deleted_count > 0:
                    self.update_treeview(self.books_manual_tree, self.books_data)
                    self.update_treeview(self.books_tree, self.books_data) # Update main import tree
                    self.update_books_stats()
                    self.log_message(f"Deleted {deleted_count} Books entries.")
                    messagebox.showinfo("Success", f"{deleted_count} entries deleted successfully.")
                else:
                    messagebox.showinfo("Info", "No matching entries found to delete.")
            
            except Exception as e:
                self.log_message(f"Error deleting Books entries: {str(e)}", error=True)
                messagebox.showerror("Error", f"Failed to delete entries: {str(e)}")

    def clear_gstr2a_manual(self):
        """Clears all manual GSTR-2A entries."""
        # This method is intended to clear ALL GSTR-2A data, regardless of source.
        # Its name "clear_gstr2a_manual" might be misleading if it clears all.
        # Renamed clear_all_gstr2a to be more explicit.
        self.clear_all_gstr2a()

    def clear_books_manual(self):
        """Clears all manual Books entries."""
        # This method is intended to clear ALL Books data, regardless of source.
        # Its name "clear_books_manual" might be misleading if it clears all.
        # Renamed clear_all_books to be more explicit.
        self.clear_all_books()

    # --- DATA MANAGEMENT METHODS ---
    def update_treeview(self, tree, df):
        """
        Updates a given Treeview widget with data from a DataFrame.
        Formats dates and numbers for display.
        """
        # Clear existing data in the treeview
        for item in tree.get_children():
            tree.delete(item)
        
        if df.empty:
            return

        # Get the columns expected by the treeview (display names)
        tree_columns_display = tree['columns']
        
        # Map display names to internal DataFrame column names
        # This mapping should be consistent with `clean_and_transform_data`
        display_to_internal_map = {
            "Invoice No": "invoice_no",
            "Invoice Date": "invoice_date",
            "Supplier GSTIN": "supplier_gstin",
            "Taxable Value": "taxable_value",
            "CGST": "cgst",
            "SGST": "sgst",
            "IGST": "igst",
            "Total Amount": "total_amount",
            "Place of Supply": "place_of_supply",
            "Book Entry Date": "book_entry_date", # Only for books
            "Source": "source", # For reconciliation results
            "Issue Type": "issue_type", # For reconciliation results
            "GSTR-2A Date": "gstr_2a_date", # For reconciliation results (internal name might differ)
            "Books Date": "books_date", # For reconciliation results (internal name might differ)
            "GSTR-2A GSTIN": "gstr_2a_gstin", # For reconciliation results
            "Books GSTIN": "books_gstin", # For reconciliation results
            "Amount Diff": "amount_diff", # For reconciliation results
            "Tax Diff": "tax_diff", # For reconciliation results
            "Details": "details" # For reconciliation results
        }

        # Insert new data row by row
        for _, row in df.iterrows():
            values = []
            for col_display_name in tree_columns_display:
                # Get the internal column name from the map, default to lowercase_underscore
                col_internal_name = display_to_internal_map.get(col_display_name, col_display_name.lower().replace(' ', '_'))
                
                if col_internal_name in row and pd.notna(row[col_internal_name]):
                    value = row[col_internal_name]
                    # Format dates
                    if 'date' in col_internal_name and isinstance(value, pd.Timestamp):
                        value = value.strftime('%d/%m/%Y')
                    # Format numbers (currency)
                    elif isinstance(value, (int, float)):
                        value = f"₹{value:,.2f}"
                    else:
                        value = str(value)
                    values.append(value)
                else:
                    values.append("") # Append empty string if column not found or value is NaN
            tree.insert("", "end", values=values)

    def update_gstr2a_stats(self):
        """Updates the statistics displayed for GSTR-2A data."""
        if self.gstr2a_data.empty:
            self.gstr2a_stats.set("Records: 0 | Total Value: ₹0 | Total Tax: ₹0")
            return
            
        total_value = self.gstr2a_data['taxable_value'].sum() if 'taxable_value' in self.gstr2a_data.columns else 0
        total_tax = 0
        if 'cgst' in self.gstr2a_data.columns: total_tax += self.gstr2a_data['cgst'].sum()
        if 'sgst' in self.gstr2a_data.columns: total_tax += self.gstr2a_data['sgst'].sum()
        if 'igst' in self.gstr2a_data.columns: total_tax += self.gstr2a_data['igst'].sum()

        self.gstr2a_stats.set(f"Records: {len(self.gstr2a_data)} | Total Value: ₹{total_value:,.2f} | Total Tax: ₹{total_tax:,.2f}")

    def update_books_stats(self):
        """Updates the statistics displayed for Books data."""
        if self.books_data.empty:
            self.books_stats.set("Records: 0 | Total Value: ₹0 | Total Tax: ₹0")
            return
            
        total_value = self.books_data['taxable_value'].sum() if 'taxable_value' in self.books_data.columns else 0
        total_tax = 0
        if 'cgst' in self.books_data.columns: total_tax += self.books_data['cgst'].sum()
        if 'sgst' in self.books_data.columns: total_tax += self.books_data['sgst'].sum()
        if 'igst' in self.books_data.columns: total_tax += self.books_data['igst'].sum()

        self.books_stats.set(f"Records: {len(self.books_data)} | Total Value: ₹{total_value:,.2f} | Total Tax: ₹{total_tax:,.2f}")

    def clear_gstr2a_import(self):
        """Clears only the imported GSTR-2A data (resets the DataFrame and UI)."""
        if messagebox.askyesno("Confirm", "Clear all imported GSTR-2A data? This will clear all GSTR-2A data, including manual entries."):
            self.gstr2a_data = pd.DataFrame(columns=[
                'invoice_no', 'invoice_date', 'supplier_gstin', 'taxable_value', 
                'cgst', 'sgst', 'igst', 'total_amount', 'place_of_supply', 'match_key'
            ])
            self.update_treeview(self.gstr2a_tree, self.gstr2a_data)
            self.update_treeview(self.gstr2a_manual_tree, self.gstr2a_data) # Clear manual tree too
            self.update_gstr2a_stats()
            self.log_message("Cleared all GSTR-2A data.")
            messagebox.showinfo("Cleared", "All GSTR-2A data has been cleared.")

    def clear_books_import(self):
        """Clears only the imported Books data (resets the DataFrame and UI)."""
        if messagebox.askyesno("Confirm", "Clear all imported Books data? This will clear all Books data, including manual entries."):
            self.books_data = pd.DataFrame(columns=[
                'invoice_no', 'invoice_date', 'supplier_gstin', 'taxable_value', 
                'cgst', 'sgst', 'igst', 'total_amount', 'place_of_supply', 
                'book_entry_date', 'match_key'
            ])
            self.update_treeview(self.books_tree, self.books_data)
            self.update_treeview(self.books_manual_tree, self.books_data) # Clear manual tree too
            self.update_books_stats()
            self.log_message("Cleared all Books data.")
            messagebox.showinfo("Cleared", "All Books data has been cleared.")

    def clear_all_gstr2a(self):
        """Clears ALL GSTR-2A data (imported and manual)."""
        self.clear_gstr2a_import() # This method now handles clearing all GSTR-2A data

    def clear_all_books(self):
        """Clears ALL Books data (imported and manual)."""
        self.clear_books_import() # This method now handles clearing all Books data

    # --- RECONCILIATION METHODS ---
    def run_reconciliation(self):
        """Initiates the reconciliation process."""
        if self.gstr2a_data.empty or self.books_data.empty:
            messagebox.showerror("Error", "Both GSTR-2A and Books data must be loaded before reconciliation.")
            self.log_message("Reconciliation aborted: Data missing.", error=True)
            return
        
        try:
            self.recon_status.set("Running...")
            self.root.update_idletasks() # Update UI to show status immediately
            self.log_message("Starting reconciliation process...")
            
            # Define essential columns for reconciliation
            # These are the internal names after clean_and_transform_data
            required_cols = ['invoice_no', 'supplier_gstin', 'invoice_date', 
                             'taxable_value', 'cgst', 'sgst', 'igst', 'total_amount', 'match_key']
            
            # Check if essential columns are present after cleaning
            gstr2a_missing_cols = [col for col in required_cols if col not in self.gstr2a_data.columns]
            books_missing_cols = [col for col in required_cols if col not in self.books_data.columns]

            if gstr2a_missing_cols:
                messagebox.showerror("Error", f"Required columns missing in GSTR-2A data: {', '.join(gstr2a_missing_cols)}. Please check your data and mapping.")
                self.recon_status.set("Error")
                self.log_message(f"Reconciliation failed: Missing required columns in GSTR-2A: {gstr2a_missing_cols}", error=True)
                return
            
            if books_missing_cols:
                messagebox.showerror("Error", f"Required columns missing in Books data: {', '.join(books_missing_cols)}. Please check your data and mapping.")
                self.recon_status.set("Error")
                self.log_message(f"Reconciliation failed: Missing required columns in Books: {books_missing_cols}", error=True)
                return

            # Perform the core reconciliation logic
            results = self.perform_reconciliation()
            self.reconciliation_results = results
            
            # Update UI with reconciliation results
            self.update_results_ui(results)
            
            # Generate and display insights
            self.generate_insights(results)
            
            self.recon_status.set("Completed")
            self.log_message("Reconciliation completed successfully.")
            messagebox.showinfo("Success", "Reconciliation completed successfully.")
            
        except Exception as e:
            self.recon_status.set("Error")
            self.log_message(f"Reconciliation error: {str(e)}", error=True)
            messagebox.showerror("Error", f"Reconciliation failed: {str(e)}")

    def perform_reconciliation(self):
        """
        Performs the core reconciliation logic between GSTR-2A and Books data.
        Identifies missing invoices, mismatches in date, amount, tax, and GSTIN,
        and detects duplicates.
        """
        results = []
        
        # Ensure 'match_key' is present and drop rows where it's NaN for reconciliation
        # Use .copy() to avoid SettingWithCopyWarning
        gstr_data_filtered = self.gstr2a_data.dropna(subset=['match_key']).copy()
        books_data_filtered = self.books_data.dropna(subset=['match_key']).copy()

        # Identify duplicates within each dataset
        # Keep all occurrences of duplicates to report them
        gstr_duplicates = gstr_data_filtered[gstr_data_filtered.duplicated(subset=['match_key'], keep=False)]
        if not gstr_duplicates.empty:
            for _, row in gstr_duplicates.iterrows():
                results.append({
                    "Invoice No": row.get('invoice_no', ''),
                    "Source": "GSTR-2A",
                    "Issue Type": "Duplicate in GSTR-2A",
                    "GSTR-2A Date": row.get('invoice_date', pd.NaT),
                    "Books Date": pd.NaT,
                    "GSTR-2A GSTIN": row.get('supplier_gstin', ''),
                    "Books GSTIN": '',
                    "Amount Diff": 0.0,
                    "Tax Diff": 0.0,
                    "Details": "Duplicate invoice found in GSTR-2A data."
                })

        books_duplicates = books_data_filtered[books_data_filtered.duplicated(subset=['match_key'], keep=False)]
        if not books_duplicates.empty:
            for _, row in books_duplicates.iterrows():
                results.append({
                    "Invoice No": row.get('invoice_no', ''),
                    "Source": "Books",
                    "Issue Type": "Duplicate in Books",
                    "GSTR-2A Date": pd.NaT,
                    "Books Date": row.get('invoice_date', pd.NaT),
                    "GSTR-2A GSTIN": '',
                    "Books GSTIN": row.get('supplier_gstin', ''),
                    "Amount Diff": 0.0,
                    "Tax Diff": 0.0,
                    "Details": "Duplicate invoice found in Books data."
                })

        # Remove duplicates from filtered dataframes for core matching logic
        # Keep 'first' occurrence of the match_key
        gstr_data_unique = gstr_data_filtered.drop_duplicates(subset=['match_key'], keep='first')
        books_data_unique = books_data_filtered.drop_duplicates(subset=['match_key'], keep='first')

        # Create sets of unique match keys for efficient lookup
        gstr_keys = set(gstr_data_unique['match_key'])
        books_keys = set(books_data_unique['match_key'])
        
        # Find missing in books (present in GSTR-2A but not in Books)
        missing_in_books_keys = gstr_keys - books_keys
        for key in missing_in_books_keys:
            row = gstr_data_unique[gstr_data_unique['match_key'] == key].iloc[0]
            results.append({
                "Invoice No": row.get('invoice_no', ''),
                "Source": "GSTR-2A",
                "Issue Type": "Missing in Books",
                "GSTR-2A Date": row.get('invoice_date', pd.NaT),
                "Books Date": pd.NaT, # Explicitly NaT for missing
                "GSTR-2A GSTIN": row.get('supplier_gstin', ''),
                "Books GSTIN": '', # Explicitly empty for missing
                "Amount Diff": row.get('total_amount', 0.0),
                "Tax Diff": (row.get('cgst', 0.0) + row.get('sgst', 0.0) + row.get('igst', 0.0)),
                "Details": "Invoice found in GSTR-2A but not in Books."
            })
        
        # Find missing in GSTR-2A (present in Books but not in GSTR-2A)
        missing_in_gstr2a_keys = books_keys - gstr_keys
        for key in missing_in_gstr2a_keys:
            row = books_data_unique[books_data_unique['match_key'] == key].iloc[0]
            results.append({
                "Invoice No": row.get('invoice_no', ''),
                "Source": "Books",
                "Issue Type": "Missing in GSTR-2A",
                "GSTR-2A Date": pd.NaT, # Explicitly NaT for missing
                "Books Date": row.get('invoice_date', pd.NaT),
                "GSTR-2A GSTIN": '', # Explicitly empty for missing
                "Books GSTIN": row.get('supplier_gstin', ''),
                "Amount Diff": -row.get('total_amount', 0.0), # Negative indicates missing from GSTR-2A perspective
                "Tax Diff": -(row.get('cgst', 0.0) + row.get('sgst', 0.0) + row.get('igst', 0.0)),
                "Details": "Invoice found in Books but not in GSTR-2A."
            })
        
        # Find matching invoices and check for mismatches
        common_keys = gstr_keys & books_keys
        for key in common_keys:
            gstr_row = gstr_data_unique[gstr_data_unique['match_key'] == key].iloc[0]
            book_row = books_data_unique[books_data_unique['match_key'] == key].iloc[0]
            
            issues = []
            
            # Date comparison
            gstr_date = gstr_row.get('invoice_date')
            book_date = book_row.get('invoice_date')
            
            if pd.notna(gstr_date) and pd.notna(book_date):
                try:
                    # Ensure both are datetime objects for comparison
                    gstr_date_dt = pd.to_datetime(gstr_date, errors='coerce')
                    book_date_dt = pd.to_datetime(book_date, errors='coerce')

                    if pd.notna(gstr_date_dt) and pd.notna(book_date_dt):
                        date_diff = (gstr_date_dt - book_date_dt).days
                        if abs(date_diff) > self.date_tolerance:
                            issues.append(f"Date diff: {abs(date_diff)} days")
                    else:
                        issues.append("Date format error in one or both records") # If conversion failed for one or both
                except Exception as e:
                    issues.append(f"Date comparison error: {e}")
            elif pd.notna(gstr_date) != pd.notna(book_date): # One date is missing, the other is present
                issues.append("One invoice date missing")
            
            # GSTIN comparison
            gstr_gstin = str(gstr_row.get('supplier_gstin', '')).strip().upper()
            book_gstin = str(book_row.get('supplier_gstin', '')).strip().upper()
            if gstr_gstin != book_gstin:
                issues.append("GSTIN mismatch")
            
            # Amount comparison
            gstr_total = gstr_row.get('total_amount', 0.0)
            book_total = book_row.get('total_amount', 0.0)
            amount_diff = gstr_total - book_total
            if abs(amount_diff) > self.amount_tolerance:
                issues.append(f"Amount diff: ₹{amount_diff:.2f}")
            
            # Tax comparison
            gstr_tax = (gstr_row.get('cgst', 0.0) + gstr_row.get('sgst', 0.0) + gstr_row.get('igst', 0.0))
            book_tax = (book_row.get('cgst', 0.0) + book_row.get('sgst', 0.0) + book_row.get('igst', 0.0))
            tax_diff = gstr_tax - book_tax
            if abs(tax_diff) > self.amount_tolerance:
                issues.append(f"Tax diff: ₹{tax_diff:.2f}")
            
            # If any issues found, add to results
            if issues:
                result = {
                    "Invoice No": gstr_row.get('invoice_no', ''), # Use GSTR-2A invoice no as primary
                    "Source": "Both",
                    "Issue Type": ", ".join(issues),
                    "GSTR-2A Date": gstr_date,
                    "Books Date": book_date,
                    "GSTR-2A GSTIN": gstr_gstin,
                    "Books GSTIN": book_gstin,
                    "Amount Diff": amount_diff,
                    "Tax Diff": tax_diff,
                    "Details": "; ".join(issues)
                }
                results.append(result)
        
        # Convert list of dicts to DataFrame
        return pd.DataFrame(results)

    def update_results_ui(self, results):
        """
        Updates the Treeviews in the Reconciliation tab with the reconciliation results.
        Populates different tabs based on discrepancy types.
        """
        # Clear all existing data in all discrepancy trees
        for tree in self.discrepancy_trees.values():
            for item in tree.get_children():
                tree.delete(item)
        
        if results.empty:
            self.log_message("No discrepancies found during reconciliation.")
            return
        
        # Populate each treeview based on issue type
        for _, row in results.iterrows():
            # Ensure all values are properly formatted strings for Treeview insertion
            gstr2a_date_str = row['GSTR-2A Date'].strftime('%d/%m/%Y') if pd.notna(row['GSTR-2A Date']) else ''
            books_date_str = row['Books Date'].strftime('%d/%m/%Y') if pd.notna(row['Books Date']) else ''

            values = [
                str(row.get("Invoice No", "")),
                str(row.get("Source", "")),
                str(row.get("Issue Type", "")),
                gstr2a_date_str,
                books_date_str,
                str(row.get("GSTR-2A GSTIN", "")),
                str(row.get("Books GSTIN", "")),
                f"₹{row.get('Amount Diff', 0.0):,.2f}",
                f"₹{row.get('Tax Diff', 0.0):,.2f}",
                str(row.get("Details", ""))
            ]
            
            # Add to the 'All Discrepancies' tab
            self.discrepancy_trees['all'].insert("", "end", values=values)
            
            # Add to specific tabs based on issue type
            issue_type_lower = row.get('Issue Type', '').lower()
            
            if "missing in books" in issue_type_lower:
                self.discrepancy_trees['missing_in_books'].insert("", "end", values=values)
            elif "missing in gstr2a" in issue_type_lower:
                self.discrepancy_trees['missing_in_gstr2a'].insert("", "end", values=values)
            # Check for "date" and "mismatch" as separate words to avoid partial matches
            elif "date" in issue_type_lower and "mismatch" in issue_type_lower:
                self.discrepancy_trees['date_mismatch'].insert("", "end", values=values)
            elif ("amount" in issue_type_lower or "tax" in issue_type_lower) and "mismatch" in issue_type_lower:
                self.discrepancy_trees['amount_mismatch'].insert("", "end", values=values)
            elif "gstin mismatch" in issue_type_lower:
                self.discrepancy_trees['gstin_mismatch'].insert("", "end", values=values)
            elif "duplicate" in issue_type_lower:
                self.discrepancy_trees['duplicates'].insert("", "end", values=values)

    # --- INSIGHTS AND REPORTING ---
    def generate_insights(self, results):
        """
        Generates visualizations and a summary report based on reconciliation results.
        """
        # Clear previous plots from all subplots
        for ax in self.axs.flatten():
            ax.clear()
        
        # Prepare and update the summary text
        summary = self.get_summary_text(results)
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(tk.END, summary)
        self.summary_text.config(state=tk.DISABLED)
        
        # Generate and draw visualizations
        self.generate_visualizations_charts(results) # Renamed to avoid confusion with overall method
        self.canvas.draw() # Redraw the matplotlib canvas

    def get_summary_text(self, results):
        """Generates a detailed text summary of the reconciliation results."""
        total_discrepancies = len(results)
        
        # Count specific issue types
        # Use .str.contains for more robust matching of issue types that might be combined
        missing_in_books = results[results['Issue Type'].str.contains('Missing in Books', case=False, na=False)].shape[0] if not results.empty else 0
        missing_in_gstr2a = results[results['Issue Type'].str.contains('Missing in GSTR-2A', case=False, na=False)].shape[0] if not results.empty else 0
        
        date_mismatch = results[results['Issue Type'].str.contains('Date diff', case=False, na=False)].shape[0] if not results.empty else 0
        amount_tax_mismatch = results[results['Issue Type'].str.contains('Amount diff|Tax diff', case=False, na=False)].shape[0] if not results.empty else 0
        gstin_mismatch = results[results['Issue Type'].str.contains('GSTIN mismatch', case=False, na=False)].shape[0] if not results.empty else 0
        duplicates = results[results['Issue Type'].str.contains('Duplicate', case=False, na=False)].shape[0] if not results.empty else 0
        
        total_amount_diff = results['Amount Diff'].sum() if 'Amount Diff' in results.columns and not results.empty else 0.0
        total_tax_diff = results['Tax Diff'].sum() if 'Tax Diff' in results.columns and not results.empty else 0.0
        
        summary = f"""--- GST Reconciliation Report ---
Generated On: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Overall Summary:
- Total Discrepancies Found: {total_discrepancies}

Breakdown by Discrepancy Type:
- Invoices Missing in Books: {missing_in_books}
- Invoices Missing in GSTR-2A: {missing_in_gstr2a}
- Date Mismatches: {date_mismatch}
- Amount/Tax Mismatches: {amount_tax_mismatch}
- GSTIN Mismatches: {gstin_mismatch}
- Duplicate Entries: {duplicates}

Financial Impact of Discrepancies:
- Total Amount Difference: ₹{total_amount_diff:,.2f}
- Total Tax Difference: ₹{total_tax_diff:,.2f}

Tolerance Settings Used:
- Date Difference Tolerance: {self.date_tolerance} days
- Amount/Tax Difference Tolerance: ₹{self.amount_tolerance:,.2f}

Recommended Action Items:
1. Review and rectify invoices missing in either GSTR-2A or Books.
2. Investigate and correct amount/tax discrepancies.
3. Verify and update incorrect GSTINs.
4. Correct invoice date mismatches.
5. Address and remove any duplicate entries in source data.
"""
        return summary

    def generate_visualizations_charts(self, results):
        """Generates various charts to visualize reconciliation insights."""
        # Discrepancy type distribution (Top-Left)
        ax1 = self.axs[0, 0]
        if not results.empty and 'Issue Type' in results.columns:
            # Count occurrences of each primary issue type
            # Use a more robust way to extract primary issue type, handling combined strings
            issue_counts = results['Issue Type'].apply(lambda x: x.split(',')[0].strip() if isinstance(x, str) else 'Unknown').value_counts().head(10)
            if not issue_counts.empty:
                issue_counts.plot(kind='bar', ax=ax1, color='skyblue')
                ax1.set_title('Top Discrepancy Types')
                ax1.set_ylabel('Count')
                ax1.tick_params(axis='x', rotation=45, labelsize=8)
                ax1.grid(axis='y', linestyle='--', alpha=0.7)
            else:
                ax1.text(0.5, 0.5, 'No discrepancy data', ha='center', va='center', transform=ax1.transAxes)
        else:
            ax1.text(0.5, 0.5, 'No discrepancy data', ha='center', va='center', transform=ax1.transAxes)
        
        # Financial impact by issue type (Top-Right)
        ax2 = self.axs[0, 1]
        if not results.empty and 'Issue Type' in results.columns and 'Amount Diff' in results.columns and 'Tax Diff' in results.columns:
            # Group by primary issue type and sum absolute differences
            financial_impact = results.groupby(results['Issue Type'].apply(lambda x: x.split(',')[0].strip() if isinstance(x, str) else 'Unknown'))[['Amount Diff', 'Tax Diff']].sum().abs()
            if not financial_impact.empty:
                financial_impact.plot(kind='bar', ax=ax2, stacked=True, color=['lightcoral', 'lightgreen'])
                ax2.set_title('Financial Impact by Issue Type')
                ax2.set_ylabel('Absolute Amount (₹)')
                ax2.tick_params(axis='x', rotation=45, labelsize=8)
                ax2.legend(['Amount Difference', 'Tax Difference'], loc='upper left', fontsize=8)
                ax2.grid(axis='y', linestyle='--', alpha=0.7)
            else:
                ax2.text(0.5, 0.5, 'No financial impact data', ha='center', va='center', transform=ax2.transAxes)
        else:
            ax2.text(0.5, 0.5, 'No financial impact data', ha='center', va='center', transform=ax2.transAxes)
        
        # Monthly discrepancy trend (Bottom-Left)
        ax3 = self.axs[1, 0]
        if not results.empty and 'GSTR-2A Date' in results.columns:
            try:
                # Ensure 'GSTR-2A Date' is datetime and filter out NaT values before resampling
                valid_dates_results = results[pd.notna(results['GSTR-2A Date'])].copy()
                valid_dates_results['GSTR-2A Date'] = pd.to_datetime(valid_dates_results['GSTR-2A Date'])

                if not valid_dates_results.empty:
                    monthly_trend = valid_dates_results.set_index('GSTR-2A Date').resample('M').size()
                    if not monthly_trend.empty:
                        monthly_trend.plot(kind='line', ax=ax3, marker='o', color='purple')
                        ax3.set_title('Monthly Discrepancy Trend (GSTR-2A Date)')
                        ax3.set_ylabel('Count')
                        ax3.set_xlabel('Month')
                        ax3.grid(True)
                        ax3.tick_params(axis='x', rotation=45, labelsize=8)
                    else:
                        ax3.text(0.5, 0.5, 'No monthly trend data', ha='center', va='center', transform=ax3.transAxes)
                else:
                    ax3.text(0.5, 0.5, 'No valid date data for trend', ha='center', va='center', transform=ax3.transAxes)
            except Exception as e:
                ax3.text(0.5, 0.5, f'Error generating trend: {e}', ha='center', va='center', transform=ax3.transAxes)
                self.log_message(f"Error generating monthly trend chart: {e}", error=True)
        else:
            ax3.text(0.5, 0.5, 'Date data unavailable or invalid', ha='center', va='center', transform=ax3.transAxes)
        
        # Top vendors with issues (Bottom-Right)
        ax4 = self.axs[1, 1]
        if not results.empty and 'GSTR-2A GSTIN' in results.columns:
            # Filter out empty/invalid GSTINs before counting
            valid_gstins = results[results['GSTR-2A GSTIN'].astype(str).str.strip() != '']['GSTR-2A GSTIN']
            top_vendors = valid_gstins.value_counts().head(10)
            if not top_vendors.empty:
                top_vendors.plot(kind='bar', ax=ax4, color='salmon')
                ax4.set_title('Top 10 Vendors with Discrepancies')
                ax4.set_ylabel('Count')
                ax4.set_xlabel('Supplier GSTIN')
                ax4.tick_params(axis='x', rotation=45, labelsize=8)
                ax4.grid(axis='y', linestyle='--', alpha=0.7)
            else:
                ax4.text(0.5, 0.5, 'No vendor data', ha='center', va='center', transform=ax4.transAxes)
        else:
            ax4.text(0.5, 0.5, 'No vendor data', ha='center', va='center', transform=ax4.transAxes)
        
        self.fig.tight_layout(pad=5.0) # Re-adjust layout after plotting

    # --- EXPORT METHODS ---
    def export_results(self):
        """Exports the reconciliation results to an Excel or CSV file."""
        if self.reconciliation_results is None or self.reconciliation_results.empty:
            messagebox.showinfo("Info", "No reconciliation results to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")],
            title="Save Reconciliation Report"
        )
        
        if not file_path:
            return # User cancelled the save dialog
        
        try:
            # Create a copy to format dates and amounts for export
            export_df = self.reconciliation_results.copy()
            
            # Format date columns to string for export
            for col in ['GSTR-2A Date', 'Books Date']:
                if col in export_df.columns:
                    export_df[col] = export_df[col].dt.strftime('%d/%m/%Y').fillna('')
            
            # Format numeric columns to 2 decimal places
            for col in ['Amount Diff', 'Tax Diff']:
                if col in export_df.columns:
                    export_df[col] = export_df[col].apply(lambda x: f"{x:.2f}")

            if file_path.lower().endswith('.csv'):
                export_df.to_csv(file_path, index=False)
            else:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    export_df.to_excel(writer, sheet_name='Discrepancies', index=False)
                    
                    # Add summary sheet
                    summary = self.get_summary_text(self.reconciliation_results)
                    # Create a DataFrame for the summary text. Ensure it's in a format pandas can write.
                    summary_df = pd.DataFrame([{"Reconciliation Summary": summary}])
                    summary_df.to_excel(writer, sheet_name='Summary', index=False, header=True)
            
            self.log_message(f"Exported results to: {file_path}")
            messagebox.showinfo("Success", f"Reconciliation results exported to:\n{file_path}")
            
        except Exception as e:
            self.log_message(f"Export error: {str(e)}", error=True)
            messagebox.showerror("Error", f"Failed to export results: {str(e)}")

    def export_all_data(self):
        """Exports all loaded GSTR-2A, Books, and reconciliation results to a single Excel file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Save All Data"
        )
        
        if not file_path:
            return # User cancelled
        
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                if not self.gstr2a_data.empty:
                    # Create a copy to format dates for export
                    gstr2a_export_df = self.gstr2a_data.copy()
                    if 'invoice_date' in gstr2a_export_df.columns:
                        gstr2a_export_df['invoice_date'] = gstr2a_export_df['invoice_date'].dt.strftime('%d/%m/%Y').fillna('')
                    gstr2a_export_df.to_excel(writer, sheet_name='GSTR-2A Data', index=False)
                else:
                    self.log_message("No GSTR-2A data to export.", error=False)
                    
                if not self.books_data.empty:
                    # Create a copy to format dates for export
                    books_export_df = self.books_data.copy()
                    if 'invoice_date' in books_export_df.columns:
                        books_export_df['invoice_date'] = books_export_df['invoice_date'].dt.strftime('%d/%m/%Y').fillna('')
                    if 'book_entry_date' in books_export_df.columns:
                        books_export_df['book_entry_date'] = books_export_df['book_entry_date'].dt.strftime('%d/%m/%Y').fillna('')
                    books_export_df.to_excel(writer, sheet_name='Books Data', index=False)
                else:
                    self.log_message("No Books data to export.", error=False)

                if self.reconciliation_results is not None and not self.reconciliation_results.empty:
                    # Create a copy to format dates and amounts for export
                    recon_export_df = self.reconciliation_results.copy()
                    for col in ['GSTR-2A Date', 'Books Date']:
                        if col in recon_export_df.columns:
                            recon_export_df[col] = recon_export_df[col].dt.strftime('%d/%m/%Y').fillna('')
                    for col in ['Amount Diff', 'Tax Diff']:
                        if col in recon_export_df.columns:
                            recon_export_df[col] = recon_export_df[col].apply(lambda x: f"{x:.2f}")
                    recon_export_df.to_excel(writer, sheet_name='Reconciliation Results', index=False)
                else:
                    self.log_message("No reconciliation results to export.", error=False)
            
            self.log_message(f"Exported all available data to: {file_path}")
            messagebox.showinfo("Success", f"All available data exported to:\n{file_path}")
            
        except Exception as e:
            self.log_message(f"Export error: {str(e)}", error=True)
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")

    # --- UTILITY METHODS ---
    def save_settings(self):
        """Saves the user-defined reconciliation settings."""
        try:
            # Validate and convert date tolerance
            date_tol_str = self.date_tol_var.get()
            if not date_tol_str.isdigit():
                raise ValueError("Date tolerance must be a whole number.")
            self.date_tolerance = max(0, int(date_tol_str))

            # Validate and convert amount tolerance
            amount_tol_str = self.amount_tol_var.get()
            try:
                self.amount_tolerance = max(0.0, float(amount_tol_str))
            except ValueError:
                raise ValueError("Amount/Tax tolerance must be a number.")
            
            self.auto_clean_gstin = bool(self.auto_clean_var.get())
            
            self.log_message("Settings saved successfully.")
            messagebox.showinfo("Success", "Settings saved successfully.")
            
        except ValueError as ve:
            self.log_message(f"Error saving settings: {str(ve)}", error=True)
            messagebox.showerror("Error", f"Invalid settings: {str(ve)}")
        except Exception as e:
            self.log_message(f"An unexpected error occurred while saving settings: {str(e)}", error=True)
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

    def show_data_summary(self):
        """Displays a summary of the loaded GSTR-2A and Books data in a new dialog."""
        summary = "=== Data Summary ===\n\n"
        
        # GSTR-2A summary
        summary += "GSTR-2A Data:\n"
        if self.gstr2a_data.empty:
            summary += "  No data loaded\n"
        else:
            summary += f"  Records: {len(self.gstr2a_data)}\n"
            # Ensure 'invoice_date' exists and is datetime type before min/max
            if 'invoice_date' in self.gstr2a_data.columns and pd.api.types.is_datetime64_any_dtype(self.gstr2a_data['invoice_date']):
                valid_dates = self.gstr2a_data['invoice_date'].dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min()
                    max_date = valid_dates.max()
                    summary += f"  Period: {min_date.strftime('%d/%m/%Y')} to {max_date.strftime('%d/%m/%Y')}\n"
                else:
                    summary += "  Period: N/A (No valid dates)\n"
            else:
                summary += "  Period: N/A (Invoice date column missing or invalid)\n"

            summary += f"  Total Taxable Value: ₹{self.gstr2a_data['taxable_value'].sum():,.2f}\n"
            total_tax_gstr2a = (self.gstr2a_data['cgst'].sum() if 'cgst' in self.gstr2a_data.columns else 0) + \
                               (self.gstr2a_data['sgst'].sum() if 'sgst' in self.gstr2a_data.columns else 0) + \
                               (self.gstr2a_data['igst'].sum() if 'igst' in self.gstr2a_data.columns else 0)
            summary += f"  Total Tax: ₹{total_tax_gstr2a:,.2f}\n"
            if 'supplier_gstin' in self.gstr2a_data.columns:
                summary += f"  Unique Suppliers: {self.gstr2a_data['supplier_gstin'].nunique()}\n"
        
        # Books summary
        summary += "\nBooks Data:\n"
        if self.books_data.empty:
            summary += "  No data loaded\n"
        else:
            summary += f"  Records: {len(self.books_data)}\n"
            # Ensure 'invoice_date' exists and is datetime type before min/max
            if 'invoice_date' in self.books_data.columns and pd.api.types.is_datetime64_any_dtype(self.books_data['invoice_date']):
                valid_dates = self.books_data['invoice_date'].dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min()
                    max_date = valid_dates.max()
                    summary += f"  Period: {min_date.strftime('%d/%m/%Y')} to {max_date.strftime('%d/%m/%Y')}\n"
                else:
                    summary += "  Period: N/A (No valid dates)\n"
            else:
                summary += "  Period: N/A (Invoice date column missing or invalid)\n"

            summary += f"  Total Taxable Value: ₹{self.books_data['taxable_value'].sum():,.2f}\n"
            total_tax_books = (self.books_data['cgst'].sum() if 'cgst' in self.books_data.columns else 0) + \
                              (self.books_data['sgst'].sum() if 'sgst' in self.books_data.columns else 0) + \
                              (self.books_data['igst'].sum() if 'igst' in self.books_data.columns else 0)
            summary += f"  Total Tax: ₹{total_tax_books:,.2f}\n"
            if 'supplier_gstin' in self.books_data.columns:
                summary += f"  Unique Suppliers: {self.books_data['supplier_gstin'].nunique()}\n"
        
        # Reconciliation summary
        summary += "\nReconciliation Status:\n"
        if self.reconciliation_results is None:
            summary += "  Not performed yet. Click 'Run Reconciliation' to see results.\n"
        else:
            summary += f"  Last Run: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            summary += f"  Discrepancies Found: {len(self.reconciliation_results)}\n"
            if not self.reconciliation_results.empty:
                summary += f"  Total Amount Discrepancy: ₹{self.reconciliation_results['Amount Diff'].sum():,.2f}\n"
                summary += f"  Total Tax Discrepancy: ₹{self.reconciliation_results['Tax Diff'].sum():,.2f}\n"
        
        # Create a new Toplevel window for the summary dialog
        summary_dialog = tk.Toplevel(self.root)
        summary_dialog.title("Data Summary")
        summary_dialog.geometry("600x500")
        
        # Add a ScrolledText widget to display the summary
        text = scrolledtext.ScrolledText(summary_dialog, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.insert(tk.END, summary)
        text.config(state=tk.DISABLED) # Make the text read-only
        
        # Add a close button
        ttk.Button(summary_dialog, text="Close", command=summary_dialog.destroy).pack(pady=10)

    def log_message(self, message, error=False):
        """
        Logs a message to the application's log text area and updates the status bar.
        Messages marked as errors are displayed in red.
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL) # Enable editing
        if error:
            self.log_text.tag_config("error", foreground="red") # Configure error tag
            self.log_text.insert(tk.END, log_entry, "error") # Insert with error tag
        else:
            self.log_text.insert(tk.END, log_entry) # Insert normal message
        self.log_text.config(state=tk.DISABLED) # Disable editing
        self.log_text.see(tk.END) # Scroll to the end of the log
        
        # Update status bar with the latest message
        self.status_var.set(message)

# Main part of the script to run the application
if __name__ == "__main__":
    root = tk.Tk() # Create the main Tkinter window
    app = GSTReconciliationApp(root) # Instantiate the application
    root.mainloop() # Start the Tkinter event loop
