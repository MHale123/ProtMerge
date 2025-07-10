"""
Excel Data Viewer for ProtMerge v1.2.0
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from pathlib import Path
import logging
import gc

# Self-contained Theme class (fallback if gui_main import fails)
class ViewerTheme:
    BG = "#1e1e1e"
    SECONDARY = "#2d2d2d"
    TERTIARY = "#3c3c3c"
    TEXT = "#ffffff"
    TEXT_MUTED = "#b0b0b0"
    CYAN = "#00bcd4"
    CYAN_HOVER = "#00acc1"
    GREEN = "#4caf50"
    RED = "#f44336"

# Self-contained ModernButton class with unique name to avoid conflicts
class ExcelViewerButton(tk.Button):
    def __init__(self, parent, text, command=None, style="primary", size="normal", **kwargs):
        colors = {
            "primary": (ViewerTheme.CYAN, "white"),
            "secondary": (ViewerTheme.TERTIARY, ViewerTheme.TEXT),
            "danger": (ViewerTheme.RED, "white"),
            "small": (ViewerTheme.TERTIARY, ViewerTheme.TEXT)
        }
        
        bg_color, fg_color = colors.get(style, colors["primary"])
        
        # Size configurations
        if size == "small":
            font_config = ("Segoe UI", 9, "bold")
            pad_config = {"padx": 10, "pady": 6}
        else:
            font_config = ("Segoe UI", 10, "bold")
            pad_config = {"padx": 15, "pady": 8}
        
        super().__init__(
            parent, text=text, command=command,
            bg=bg_color, fg=fg_color,
            font=font_config,
            relief="flat", borderwidth=0,
            cursor="hand2",
            activebackground=bg_color,  # Simplified hover
            activeforeground=fg_color,
            **pad_config,
            **kwargs
        )

# Try to import from gui_main safely without overriding classes
try:
    import sys
    import os
    
    # Add current directory to path if needed
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    # Import with different names to avoid conflicts
    from gui_main import Theme as MainTheme, ModernButton as MainModernButton
    
    # Use main GUI theme and buttons
    Theme = MainTheme
    ModernButton = MainModernButton
    
except ImportError as e:
    # Use fallback classes with no conflicts
    Theme = ViewerTheme
    ModernButton = ExcelViewerButton


class ExcelDataViewer:
    """Professional Excel Data Viewer with full functionality and proper file handle management"""
    
    def __init__(self, parent, file_path, close_callback=None):
        self.parent = parent
        self.file_path = Path(file_path)
        self.close_callback = close_callback
        self.logger = logging.getLogger(__name__)
        
        # File handle management
        self.excel_file = None
        self.sheet_data_cache = {}  # Cache loaded sheet data
        self.is_file_open = False
        
        # Window setup
        self.window = None
        self.notebook = None
        self.status_label = None
        self.sheet_info = None
        
        # Initialize the viewer
        self._setup_window()
        self._create_interface()
        self._load_excel_file()
        
        self.logger.info(f"Excel viewer opened for: {self.file_path.name}")
    
    def _setup_window(self):
        """Setup the main viewer window"""
        try:
            self.window = tk.Toplevel(self.parent)
            self.window.title(f"Excel Data Viewer - {self.file_path.name}")
            self.window.geometry("1400x900")
            self.window.configure(bg=Theme.BG)
            self.window.transient(self.parent)
            self.window.grab_set()
            self.window.protocol("WM_DELETE_WINDOW", self._close_viewer)
        except Exception as e:
            self.logger.error(f"Failed to setup window: {e}")
            raise
    
    def _create_interface(self):
        """Create the main viewer interface"""
        try:
            # Header with file info and controls
            self._create_header()
            
            # Main content area with notebook
            self._create_content_area()
            
            # Status bar
            self._create_status_bar()
        except Exception as e:
            self.logger.error(f"Failed to create interface: {e}")
            raise
    
    def _create_header(self):
        """Create header with file info and controls"""
        header_frame = tk.Frame(self.window, bg=Theme.CYAN, height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Left side - File info
        info_frame = tk.Frame(header_frame, bg=Theme.CYAN)
        info_frame.pack(side=tk.LEFT, fill=tk.Y, padx=20, pady=15)
        
        tk.Label(info_frame, text="📊 Excel Data Viewer",
                font=("Segoe UI", 16, "bold"),
                fg="white", bg=Theme.CYAN).pack(anchor=tk.W)
        
        tk.Label(info_frame, text=f"File: {self.file_path.name}",
                font=("Segoe UI", 11),
                fg="white", bg=Theme.CYAN).pack(anchor=tk.W, pady=(2, 0))
        
        # Right side - Controls
        controls_frame = tk.Frame(header_frame, bg=Theme.CYAN)
        controls_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=20, pady=15)
        
        btn_frame = tk.Frame(controls_frame, bg=Theme.CYAN)
        btn_frame.pack(side=tk.RIGHT)
        
        ModernButton(btn_frame, "🔍 Search", self._show_search_dialog, "secondary").pack(side=tk.LEFT, padx=(0, 10))
        ModernButton(btn_frame, "💾 Export Sheet", self._export_current_sheet, "secondary").pack(side=tk.LEFT, padx=(0, 10))
        ModernButton(btn_frame, "🔄 Refresh", self._refresh_all, "secondary").pack(side=tk.LEFT, padx=(0, 10))
        ModernButton(btn_frame, "✕ Close", self._close_viewer, "danger").pack(side=tk.LEFT)
    
    def _create_content_area(self):
        """Create main content area with notebook"""
        content_frame = tk.Frame(self.window, bg=Theme.BG)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create notebook for sheets
        self.notebook = ttk.Notebook(content_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Bind tab selection events
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
    
    def _create_status_bar(self):
        """Create bottom status bar"""
        status_frame = tk.Frame(self.window, bg=Theme.SECONDARY, height=30)
        status_frame.pack(fill=tk.X)
        status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(status_frame, text="Loading...", 
                                    font=("Segoe UI", 9),
                                    fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY)
        self.status_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.sheet_info = tk.Label(status_frame, text="", 
                                  font=("Segoe UI", 9),
                                  fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY)
        self.sheet_info.pack(side=tk.RIGHT, padx=10, pady=5)
    
    def _load_excel_file(self):
        """Load all sheets from the Excel file with proper file handle management"""
        try:
            self.status_label.config(text="Loading Excel file...")
            self.window.update()
            
            # Open Excel file with proper context management
            try:
                self.excel_file = pd.ExcelFile(self.file_path)
                self.is_file_open = True
                total_sheets = len(self.excel_file.sheet_names)
                
                self.logger.info(f"Loading {total_sheets} sheets from {self.file_path.name}")
                
                for i, sheet_name in enumerate(self.excel_file.sheet_names):
                    self.status_label.config(text=f"Loading sheet {i+1}/{total_sheets}: {sheet_name}")
                    self.window.update()
                    
                    try:
                        # Read and cache the sheet data
                        df = self._read_sheet_safely(sheet_name)
                        self.sheet_data_cache[sheet_name] = df
                        self._create_sheet_tab(sheet_name, df)
                        
                    except Exception as e:
                        self.logger.error(f"Error loading sheet '{sheet_name}': {e}")
                        self._create_error_tab(sheet_name, str(e))
                
                # Close the Excel file handle immediately after loading all data
                self._close_file_handle()
                
                # Select first tab if available
                if self.notebook.tabs():
                    self.notebook.select(0)
                    self._update_status()
                
                self.status_label.config(text="Ready - File handles closed")
                
            except Exception as e:
                self.logger.error(f"Failed to open Excel file: {e}")
                raise
                
        except Exception as e:
            self.logger.error(f"Failed to load Excel file: {e}")
            messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")
            self._close_viewer()
    
    def _read_sheet_safely(self, sheet_name):
        """Safely read a sheet with proper error handling"""
        try:
            # Use the open Excel file handle
            if self.excel_file is not None:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
            else:
                # Fallback: read directly from file (less efficient but safer)
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error reading sheet '{sheet_name}': {e}")
            # Return empty DataFrame on error
            return pd.DataFrame()
    
    def _close_file_handle(self):
        """Safely close the Excel file handle"""
        try:
            if self.excel_file is not None:
                # Close the ExcelFile object
                if hasattr(self.excel_file, 'close'):
                    self.excel_file.close()
                elif hasattr(self.excel_file, 'book'):
                    # For xlrd-based ExcelFile objects
                    if hasattr(self.excel_file.book, 'release_resources'):
                        self.excel_file.book.release_resources()
                
                self.excel_file = None
                self.is_file_open = False
                self.logger.info(f"Closed file handle for: {self.file_path.name}")
                
        except Exception as e:
            self.logger.warning(f"Error closing file handle: {e}")
            # Force cleanup
            self.excel_file = None
            self.is_file_open = False
    
    def _create_sheet_tab(self, sheet_name, df):
        """Create a tab for a sheet with full data display"""
        try:
            # Create main frame for this sheet
            main_frame = tk.Frame(self.notebook, bg=Theme.BG)
            tab_text = f"{sheet_name} ({len(df)} rows)" if not df.empty else f"{sheet_name} (empty)"
            self.notebook.add(main_frame, text=tab_text)
            
            # Store sheet data reference (data is already cached)
            main_frame.sheet_name = sheet_name
            main_frame.df = df  # Reference to cached data
            
            if df.empty:
                self._create_empty_sheet_display(main_frame, sheet_name)
                return
            
            # Sheet header with controls
            self._create_sheet_header(main_frame, sheet_name, df)
            
            # Data display area
            data_frame = tk.Frame(main_frame, bg=Theme.BG)
            data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Store reference for refresh
            main_frame.data_frame = data_frame
            
            # Create initial data table
            self._create_data_table(data_frame, df, 100)  # Default 100 rows
            
        except Exception as e:
            self.logger.error(f"Error creating sheet tab for '{sheet_name}': {e}")
    
    def _create_sheet_header(self, main_frame, sheet_name, df):
        """Create header controls for a sheet"""
        header_frame = tk.Frame(main_frame, bg=Theme.SECONDARY, height=50)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Left side - Sheet info
        info_frame = tk.Frame(header_frame, bg=Theme.SECONDARY)
        info_frame.pack(side=tk.LEFT, fill=tk.Y, padx=15, pady=10)
        
        tk.Label(info_frame, text=f"📋 {sheet_name}", 
                font=("Segoe UI", 12, "bold"),
                fg=Theme.TEXT, bg=Theme.SECONDARY).pack(side=tk.LEFT)
        
        tk.Label(info_frame, text=f"({len(df)} rows × {len(df.columns)} columns)", 
                font=("Segoe UI", 10),
                fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY).pack(side=tk.LEFT, padx=(10, 0))
        
        # Right side - View controls
        controls_frame = tk.Frame(header_frame, bg=Theme.SECONDARY)
        controls_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=15, pady=8)
        
        # Rows per page selector
        tk.Label(controls_frame, text="Rows:", 
                font=("Segoe UI", 9),
                fg=Theme.TEXT, bg=Theme.SECONDARY).pack(side=tk.LEFT, padx=(0, 5))
        
        rows_var = tk.StringVar(value="100")
        rows_combo = ttk.Combobox(controls_frame, textvariable=rows_var, 
                                 values=["50", "100", "200", "500", "1000", "All"], 
                                 width=8, state="readonly")
        rows_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        # Store references
        main_frame.rows_var = rows_var
        main_frame.rows_combo = rows_combo
        
        # Bind change event
        rows_combo.bind("<<ComboboxSelected>>", 
                       lambda e: self._refresh_sheet_data(main_frame))
        
        # Export this sheet button
        ModernButton(controls_frame, "💾", 
                    lambda: self._export_sheet(sheet_name, df), 
                    "secondary", "small").pack(side=tk.LEFT)
    
    def _create_data_table(self, parent, df, max_rows=None):
        """Create the data table with Excel-like functionality"""
        try:
            # Clear existing content
            for widget in parent.winfo_children():
                widget.destroy()
            
            # Determine rows to display
            total_rows = len(df)
            if max_rows is None or max_rows >= total_rows:
                display_df = df
                showing_text = f"Showing all {total_rows} rows"
            else:
                display_df = df.head(max_rows)
                showing_text = f"Showing {max_rows} of {total_rows} rows"
            
            # Info label
            info_frame = tk.Frame(parent, bg=Theme.BG)
            info_frame.pack(fill=tk.X, pady=(0, 5))
            
            tk.Label(info_frame, text=showing_text,
                    font=("Segoe UI", 10),
                    fg=Theme.TEXT_MUTED, bg=Theme.BG).pack(side=tk.LEFT)
            
            # Add column count
            tk.Label(info_frame, text=f"• {len(df.columns)} columns",
                    font=("Segoe UI", 10),
                    fg=Theme.TEXT_MUTED, bg=Theme.BG).pack(side=tk.LEFT, padx=(10, 0))
            
            # Create scrollable frame for the table
            table_frame = tk.Frame(parent, bg=Theme.BG)
            table_frame.pack(fill=tk.BOTH, expand=True)
            
            # Create treeview with all columns
            columns = [str(col) for col in display_df.columns]
            tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
            
            # Configure columns
            for col in columns:
                tree.heading(col, text=col)
                # Smart column width calculation
                header_width = len(col) * 8
                if len(display_df) > 0:
                    sample_values = display_df[col].astype(str).head(5)
                    content_width = max(len(str(val)) for val in sample_values) * 8 if len(sample_values) > 0 else 80
                else:
                    content_width = 80
                optimal_width = max(min(max(header_width, content_width), 300), 80)
                tree.column(col, width=optimal_width, anchor=tk.W)
            
            # Add scrollbars
            v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
            h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=tree.xview)
            tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            # Grid layout
            tree.grid(row=0, column=0, sticky='nsew')
            v_scrollbar.grid(row=0, column=1, sticky='ns')
            h_scrollbar.grid(row=1, column=0, sticky='ew')
            
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)
            
            # Style the treeview
            style = ttk.Style()
            style.configure('Treeview', 
                           background=Theme.TERTIARY, 
                           foreground=Theme.TEXT, 
                           fieldbackground=Theme.TERTIARY,
                           font=("Consolas", 9))
            style.configure('Treeview.Heading', 
                           background=Theme.SECONDARY, 
                           foreground=Theme.TEXT,
                           font=("Segoe UI", 9, "bold"))
            style.map('Treeview', background=[('selected', Theme.CYAN)])
            
            # Populate with data
            for idx, row in display_df.iterrows():
                values = []
                for col in columns:
                    val = row[col]
                    if pd.isna(val):
                        values.append("")
                    elif isinstance(val, float):
                        if val.is_integer():
                            values.append(str(int(val)))
                        else:
                            values.append(f"{val:.6g}")
                    else:
                        values.append(str(val)[:100])  # Limit cell display length
                
                tree.insert('', 'end', values=values)
            
            # Add context menu
            self._add_context_menu(tree)
            
            # Store tree reference
            parent.tree = tree
            
            return tree
            
        except Exception as e:
            self.logger.error(f"Error creating data table: {e}")
    
    def _add_context_menu(self, tree):
        """Add right-click context menu to tree"""
        def show_context_menu(event):
            try:
                # Select item under cursor
                item = tree.identify_row(event.y)
                if item:
                    tree.selection_set(item)
                    
                    menu = tk.Menu(tree, tearoff=0, 
                                  bg=Theme.SECONDARY, fg=Theme.TEXT,
                                  activebackground=Theme.CYAN, activeforeground="white")
                    menu.add_command(label="Copy Cell", command=lambda: self._copy_cell(tree, event))
                    menu.add_command(label="Copy Row", command=lambda: self._copy_row(tree))
                    menu.add_separator()
                    menu.add_command(label="View Cell Details", command=lambda: self._view_cell_details(tree, event))
                    
                    menu.tk_popup(event.x_root, event.y_root)
            except Exception as e:
                self.logger.debug(f"Context menu error: {e}")
        
        tree.bind("<Button-3>", show_context_menu)  # Right click
    
    def _create_empty_sheet_display(self, main_frame, sheet_name):
        """Create display for empty sheets"""
        empty_frame = tk.Frame(main_frame, bg=Theme.BG)
        empty_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(empty_frame, text="📄", 
                font=("Segoe UI", 48),
                fg=Theme.TEXT_MUTED, bg=Theme.BG).pack(expand=True, pady=(100, 20))
        
        tk.Label(empty_frame, text=f"Sheet '{sheet_name}' is empty", 
                font=("Segoe UI", 14),
                fg=Theme.TEXT_MUTED, bg=Theme.BG).pack()
        
        tk.Label(empty_frame, text="No data to display", 
                font=("Segoe UI", 10),
                fg=Theme.TEXT_MUTED, bg=Theme.BG).pack(pady=(5, 0))
    
    def _create_error_tab(self, sheet_name, error_msg):
        """Create error tab for sheets that failed to load"""
        error_frame = tk.Frame(self.notebook, bg=Theme.BG)
        self.notebook.add(error_frame, text=f"{sheet_name} (Error)")
        
        tk.Label(error_frame, text="❌", 
                font=("Segoe UI", 48),
                fg=Theme.RED, bg=Theme.BG).pack(expand=True, pady=(100, 20))
        
        tk.Label(error_frame, text="Error Loading Sheet", 
                font=("Segoe UI", 14, "bold"),
                fg=Theme.RED, bg=Theme.BG).pack(pady=(0, 10))
        
        tk.Label(error_frame, text=f"Sheet: {sheet_name}", 
                font=("Segoe UI", 12),
                fg=Theme.TEXT, bg=Theme.BG).pack(pady=5)
        
        tk.Label(error_frame, text=f"Error: {error_msg}", 
                font=("Segoe UI", 10),
                fg=Theme.TEXT_MUTED, bg=Theme.BG, 
                wraplength=400, justify=tk.CENTER).pack(pady=5)
    
    # Event handlers and utility methods
    
    def _on_tab_changed(self, event=None):
        """Handle tab change events"""
        self._update_status()
    
    def _update_status(self):
        """Update status bar information"""
        try:
            current_tab = self.notebook.select()
            if current_tab:
                tab_text = self.notebook.tab(current_tab, "text")
                self.sheet_info.config(text=f"Sheet: {tab_text}")
        except Exception as e:
            self.logger.debug(f"Status update error: {e}")
    
    def _refresh_sheet_data(self, main_frame):
        """Refresh data display for a specific sheet"""
        try:
            rows_setting = main_frame.rows_var.get()
            max_rows = None if rows_setting == "All" else int(rows_setting)
            self._create_data_table(main_frame.data_frame, main_frame.df, max_rows)
        except Exception as e:
            self.logger.error(f"Refresh error: {e}")
    
    def _refresh_all(self):
        """Refresh entire viewer with proper file handle management"""
        try:
            # Clear cached data
            self.sheet_data_cache.clear()
            
            # Clear all tabs
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)
            
            # Force garbage collection to help release any lingering file handles
            gc.collect()
            
            # Reload file
            self._load_excel_file()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh:\n{e}")
    
    def _show_search_dialog(self):
        """Show search dialog"""
        search_text = simpledialog.askstring(
            "Search Data", 
            "Enter text to search for:",
            parent=self.window
        )
        
        if search_text:
            self._perform_search(search_text)
    
    def _perform_search(self, search_text):
        """Perform search across current sheet using cached data"""
        try:
            current_tab = self.notebook.select()
            if not current_tab:
                return
            
            # Get current sheet data from cache
            tab_widget = self.notebook.nametowidget(current_tab)
            if hasattr(tab_widget, 'sheet_name'):
                sheet_name = tab_widget.sheet_name
                
                # Use cached data for search (no file access needed)
                if sheet_name in self.sheet_data_cache:
                    df = self.sheet_data_cache[sheet_name]
                    
                    # Simple text search across all columns
                    mask = df.astype(str).apply(lambda x: x.str.contains(search_text, case=False, na=False)).any(axis=1)
                    results = df[mask]
                    
                    if len(results) > 0:
                        messagebox.showinfo(
                            "Search Results", 
                            f"Found {len(results)} rows containing '{search_text}'\n\n"
                            f"Advanced search and highlighting coming soon!"
                        )
                    else:
                        messagebox.showinfo("Search Results", f"No results found for '{search_text}'")
                else:
                    messagebox.showwarning("Error", "Sheet data not available for search")
            
        except Exception as e:
            messagebox.showerror("Search Error", f"Search failed:\n{e}")
    
    def _export_current_sheet(self):
        """Export currently selected sheet using cached data"""
        try:
            current_tab = self.notebook.select()
            if not current_tab:
                messagebox.showwarning("No Sheet", "No sheet selected")
                return
            
            tab_widget = self.notebook.nametowidget(current_tab)
            if hasattr(tab_widget, 'sheet_name'):
                sheet_name = tab_widget.sheet_name
                
                # Get data from cache
                if sheet_name in self.sheet_data_cache:
                    df = self.sheet_data_cache[sheet_name]
                    self._export_sheet(sheet_name, df)
                else:
                    messagebox.showwarning("Error", "Sheet data not available for export")
            else:
                messagebox.showwarning("Error", "Cannot export this sheet")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{e}")
    
    def _export_sheet(self, sheet_name, df):
        """Export specific sheet to file using cached data"""
        try:
            file_path = filedialog.asksaveasfilename(
                title=f"Export {sheet_name}",
                defaultextension=".csv",
                filetypes=[
                    ("CSV files", "*.csv"),
                    ("Excel files", "*.xlsx"),
                    ("Tab-separated", "*.tsv"),
                    ("All files", "*.*")
                ],
                parent=self.window
            )
            
            if file_path:
                # Use cached data for export (no file handle needed)
                if file_path.lower().endswith('.csv'):
                    df.to_csv(file_path, index=False)
                elif file_path.lower().endswith('.tsv'):
                    df.to_csv(file_path, index=False, sep='\t')
                else:  # Default to Excel
                    # For Excel export, we create a new file so no handle conflict
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                
                messagebox.showinfo("Success", f"Sheet exported to:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export sheet:\n{e}")
    
    def _copy_row(self, tree):
        """Copy selected row to clipboard"""
        try:
            selection = tree.selection()
            if selection:
                item = selection[0]
                values = tree.item(item)['values']
                row_text = '\t'.join(str(val) for val in values)
                self.window.clipboard_clear()
                self.window.clipboard_append(row_text)
                self.status_label.config(text="Row copied to clipboard")
        except Exception as e:
            self.logger.debug(f"Copy row error: {e}")
    
    def _copy_cell(self, tree, event):
        """Copy clicked cell to clipboard"""
        try:
            # Identify which column was clicked
            column = tree.identify_column(event.x)
            item = tree.identify_row(event.y)
            
            if item and column:
                col_index = int(column.replace('#', '')) - 1
                values = tree.item(item)['values']
                if col_index < len(values):
                    cell_value = str(values[col_index])
                    self.window.clipboard_clear()
                    self.window.clipboard_append(cell_value)
                    self.status_label.config(text="Cell copied to clipboard")
        except Exception as e:
            self.logger.debug(f"Copy cell error: {e}")
    
    def _view_cell_details(self, tree, event):
        """Show detailed view of cell contents"""
        try:
            column = tree.identify_column(event.x)
            item = tree.identify_row(event.y)
            
            if item and column:
                col_index = int(column.replace('#', '')) - 1
                values = tree.item(item)['values']
                headers = [tree.heading(col)['text'] for col in tree['columns']]
                
                if col_index < len(values) and col_index < len(headers):
                    cell_value = str(values[col_index])
                    column_name = headers[col_index]
                    
                    # Create detail dialog
                    detail_window = tk.Toplevel(self.window)
                    detail_window.title("Cell Details")
                    detail_window.geometry("400x300")
                    detail_window.configure(bg=Theme.BG)
                    detail_window.transient(self.window)
                    
                    tk.Label(detail_window, text=f"Column: {column_name}",
                            font=("Segoe UI", 12, "bold"),
                            fg=Theme.TEXT, bg=Theme.BG).pack(pady=10)
                    
                    text_widget = tk.Text(detail_window, wrap=tk.WORD, 
                                         bg=Theme.TERTIARY, fg=Theme.TEXT,
                                         font=("Consolas", 10))
                    text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                    text_widget.insert('1.0', cell_value)
                    text_widget.config(state='disabled')
                    
                    ModernButton(detail_window, "Close", detail_window.destroy).pack(pady=10)
                    
        except Exception as e:
            self.logger.debug(f"Cell details error: {e}")
    
    def _close_viewer(self):
        """Close the viewer window with proper cleanup"""
        try:
            # Close any open file handles
            self._close_file_handle()
            
            # Clear cached data
            self.sheet_data_cache.clear()
            
            # Force garbage collection
            gc.collect()
            
            # Release window grab and destroy
            if self.window:
                self.window.grab_release()
                self.window.destroy()
            
            self.logger.info(f"Excel viewer closed and file handles released for: {self.file_path.name}")
            
            if self.close_callback:
                self.close_callback()
                
        except Exception as e:
            self.logger.debug(f"Close error: {e}")
            # Force cleanup even if there are errors
            try:
                self.excel_file = None
                self.sheet_data_cache.clear()
                if self.window:
                    self.window.destroy()
            except:
                pass


# Convenience function to launch the viewer
def launch_excel_viewer(parent, file_path, close_callback=None):
    """
    Launch the Excel Data Viewer
    
    Args:
        parent: Parent window
        file_path: Path to Excel file
        close_callback: Optional callback when viewer closes
    
    Returns:
        ExcelDataViewer instance
    """
    try:
        return ExcelDataViewer(parent, file_path, close_callback)
    except Exception as e:
        error_msg = f"Failed to open Excel viewer: {e}"
        print(f"Error: {error_msg}")  # Console output for debugging
        messagebox.showerror("Error", error_msg)
        return None


# Example usage and testing
if __name__ == "__main__":
    # Set up basic logging for testing
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    
    root = tk.Tk()
    root.withdraw()  # Hide root window
    
    # Test with file dialog
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    if file_path:
        print(f"Opening Excel viewer for: {file_path}")
        viewer = launch_excel_viewer(root, file_path)
        if viewer:
            print("Excel viewer launched successfully")
            root.mainloop()
        else:
            print("Failed to launch Excel viewer")
    else:
        print("No file selected")