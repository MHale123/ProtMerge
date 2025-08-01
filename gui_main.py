"""
gui_main.py - Complete Rewrite for ProtMerge v1.2.0 (with CompletionDialog fix)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import pandas as pd
from pathlib import Path
import logging
import os
import subprocess
import time

# Assuming these modules exist in the same directory or are in PYTHONPATH
# from excel_data_viewer import launch_excel_viewer
# from config import SUPPORTED_EXCEL_FORMATS, ERROR_MESSAGES
# from similarity_gui import launch_similarity_analysis


# Mock/Placeholder for external modules for standalone rewrite
try:
    from excel_data_viewer import launch_excel_viewer
except ImportError:
    print("Warning: excel_data_viewer.py not found. Using a mock function.")
    def launch_excel_viewer(parent, file_path, on_close_callback):
        messagebox.showinfo("Data Viewer", f"Mock Data Viewer for: {file_path}")
        if on_close_callback:
            on_close_callback()
        return True

try:
    from config import SUPPORTED_EXCEL_FORMATS, ERROR_MESSAGES
except ImportError:
    print("Warning: config.py not found. Using default configurations.")
    SUPPORTED_EXCEL_FORMATS = [("Excel files", "*.xlsx *.xls")]
    ERROR_MESSAGES = {
        "FILE_LOAD_ERROR": "Failed to load file.",
        "INVALID_PROTMERGE_FILE": "This doesn't appear to be a ProtMerge results file.",
        "SHEET_NOT_FOUND": "Expected sheet 'ProtMerge_Results' not found.",
        "REQUIRED_COLUMN_MISSING": "Required column 'UniProt ID' not found in results file.",
        "INSUFFICIENT_DATA": "Need at least 2 proteins for similarity analysis.",
        "MODULE_MISSING": "Required module not found.",
        "ANALYSIS_FAILED": "Analysis failed.",
        "FILE_OPEN_ERROR": "Could not open file:",
        "SIMILARITY_ANALYSIS_FAILED": "Similarity analysis failed.",
    }

try:
    from similarity_gui import launch_similarity_analysis
except ImportError:
    print("Warning: similarity_gui.py not found. Using a mock function.")
    def launch_similarity_analysis(parent, file_path, app):
        messagebox.showinfo("Similarity Analyzer", f"Mock Similarity Analyzer for: {file_path}")
        return "complete" # Simulate completion


class Theme:
    """Modern dark theme constants"""
    BG = "#1e1e1e"
    SECONDARY = "#2d2d2d"
    TERTIARY = "#3c3c3c"
    TEXT = "#ffffff"
    TEXT_MUTED = "#b0b0b0"
    CYAN = "#00bcd4"
    CYAN_HOVER = "#00acc1"
    GREEN = "#4caf50"
    RED = "#f44336"
    PURPLE = "#9c27b0"
    ORANGE = "#ff9800"

class ModernButton(tk.Button):
    """Sleek modern button with hover animations and proper active state handling"""

    def __init__(self, parent, text, command=None, style="primary", size="normal", icon="", **kwargs):
        self.style_configs = {
            "primary": {
                "bg": Theme.CYAN,
                "hover": Theme.CYAN_HOVER,
                "fg": "white",
                "active": "#00acc1"
            },
            "secondary": {
                "bg": Theme.TERTIARY,
                "hover": "#4a4a4a",
                "fg": Theme.TEXT,
                "active": Theme.SECONDARY
            },
            "success": {
                "bg": Theme.GREEN,
                "hover": "#45a049",
                "fg": "white",
                "active": "#4caf50"
            },
            "danger": {
                "bg": Theme.RED,
                "hover": "#da190b",
                "fg": "white",
                "active": "#f44336"
            },
            "ghost": {
                "bg": Theme.SECONDARY,
                "hover": Theme.TERTIARY,
                "fg": Theme.TEXT_MUTED,
                "active": Theme.TERTIARY
            }
        }

        self.size_configs = {
            "small": {"font": ("Segoe UI", 9, "bold"), "padx": 12, "pady": 6},
            "normal": {"font": ("Segoe UI", 11, "bold"), "padx": 20, "pady": 10},
            "large": {"font": ("Segoe UI", 13, "bold"), "padx": 28, "pady": 14},
            "xl": {"font": ("Segoe UI", 15, "bold"), "padx": 36, "pady": 18}
        }

        self.current_style = style
        self.config_style = self.style_configs.get(style, self.style_configs["primary"])
        self.size_config = self.size_configs.get(size, self.size_configs["normal"])

        # Track if this is an active navigation button
        self.is_nav_active = False

        # Add icon to text if provided
        display_text = f"{icon} {text}" if icon else text

        button_kwargs = {
            "bg": self.config_style["bg"],
            "fg": self.config_style["fg"],
            "font": self.size_config["font"],
            "relief": "flat",
            "borderwidth": 0,
            "padx": self.size_config["padx"],
            "pady": self.size_config["pady"],
            "cursor": "hand2",
            "activebackground": self.config_style["active"],
            "activeforeground": self.config_style["fg"]
        }

        # Add any additional kwargs, allowing them to override defaults
        button_kwargs.update(kwargs)

        super().__init__(
            parent,
            text=display_text,
            command=command,
            **button_kwargs
        )

        self.bind("<Enter>", self._on_hover)
        self.bind("<Leave>", self._on_leave)

    def _on_hover(self, e):
        """Handle mouse enter - respect active navigation state"""
        if self['state'] != 'disabled':
            if self.is_nav_active:
                # Active nav button: use slightly darker hover
                self.config(bg=self.style_configs["primary"]["hover"])
            else:
                # Regular button or inactive nav: normal hover
                self.config(bg=self.config_style["hover"])

    def _on_leave(self, e):
        """Handle mouse leave - restore correct state"""
        if self['state'] != 'disabled':
            if self.is_nav_active:
                # Active nav button: return to primary color
                self.config(bg=self.style_configs["primary"]["bg"])
            else:
                # Regular button or inactive nav: return to original color
                self.config(bg=self.config_style["bg"])

    def set_nav_active(self, active=True):
        """Set navigation button as active/inactive"""
        self.is_nav_active = active

        if active:
            # Active navigation button: use primary styling
            primary_style = self.style_configs["primary"]
            self.config_style = primary_style
            self.config(
                bg=primary_style["bg"],
                fg=primary_style["fg"]
            )
        else:
            # Inactive navigation button: use ghost styling
            ghost_style = self.style_configs["ghost"]
            self.config_style = ghost_style
            self.config(
                bg=ghost_style["bg"],
                fg=ghost_style["fg"]
            )

class ProgressBar(tk.Frame):
    """Custom progress bar with dark theme"""

    def __init__(self, parent):
        super().__init__(parent, bg=Theme.BG)
        self.progress_value = 0

        # Use ttk.Progressbar with custom styling
        style = ttk.Style()
        style.theme_use('default')

        style.configure('Dark.Horizontal.TProgressbar',
                       background=Theme.CYAN,
                       troughcolor=Theme.TERTIARY,
                       borderwidth=1,
                       lightcolor=Theme.CYAN,
                       darkcolor=Theme.CYAN,
                       relief='flat')

        self.progress_bar = ttk.Progressbar(
            self,
            mode='determinate',
            length=500,
            style='Dark.Horizontal.TProgressbar'
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(10, 5))

        self.label = tk.Label(self, text="0%", fg=Theme.TEXT, bg=Theme.BG,
                             font=("Segoe UI", 10, "bold"))
        self.label.pack(anchor=tk.E, padx=10, pady=(2, 10))

    def set_progress(self, value):
        self.progress_value = max(0, min(100, value))
        self.progress_bar['value'] = self.progress_value
        self.label.config(text=f"{self.progress_value:.0f}%")

        self.progress_bar.update_idletasks()
        self.label.update_idletasks()
        self.update_idletasks()


class CompletionDialog:
    """Enhanced completion dialog with multiple options and scrollable content"""

    def __init__(self, parent, output_file, analysis_summary):
        self.parent = parent
        self.output_file = output_file
        self.analysis_summary = analysis_summary
        self.result = None
        self.modal = None
        self.summary_canvas = None # To hold the canvas for the summary
        self.summary_frame_inner = None # To hold the frame inside the canvas

    def show(self):
        self.modal = tk.Toplevel(self.parent)
        self.modal.title("Analysis Complete")
        self.modal.geometry("700x600") # Maintain initial size, content will manage itself
        self.modal.configure(bg=Theme.BG)
        self.modal.resizable(False, False) # Keep fixed for consistent layout
        self.modal.transient(self.parent)
        self.modal.grab_set()

        # Center on parent
        self.modal.update_idletasks()
        x = (self.parent.winfo_x() + self.parent.winfo_width() // 2) - (self.modal.winfo_width() // 2)
        y = (self.parent.winfo_y() + self.parent.winfo_height() // 2) - (self.modal.winfo_height() // 2)
        self.modal.geometry(f"+{x}+{y}")

        self._create_content()
        self.modal.focus_set()
        self.modal.wait_visibility()

        # Ensure content is correctly sized after everything is visible
        self.modal.update_idletasks()
        self._on_modal_resize(None) # Force initial resize call

        self.parent.wait_window(self.modal)
        return self.result

    def _create_content(self):
        # Header (fixed at top)
        header = tk.Frame(self.modal, bg=Theme.GREEN, height=80)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        header_content = tk.Frame(header, bg=Theme.GREEN)
        header_content.pack(expand=True, fill=tk.BOTH)

        tk.Label(header_content, text="✅ Analysis Complete",
                font=("Segoe UI", 18, "bold"), fg="white", bg=Theme.GREEN).pack(pady=(15, 5))

        tk.Label(header_content, text="Your protein analysis has finished successfully",
                font=("Segoe UI", 11), fg="white", bg=Theme.GREEN).pack()

        # Bottom buttons (fixed at bottom)
        self._create_bottom_buttons() # Call this first so it reserves space at the bottom

        # Main content area (between header and bottom buttons)
        main_content_area = tk.Frame(self.modal, bg=Theme.BG)
        main_content_area.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

        # Results summary (now scrollable)
        summary_container = self._create_results_summary(main_content_area)
        # summary_container packs with fill=tk.X and doesn't expand vertically
        # The canvas inside it will manage its own scrolling if content exceeds.
        summary_container.pack(fill=tk.X, pady=(0, 20))

        # Action options (fixed grid)
        action_options_frame = self._create_action_options(main_content_area)
        # This frame must fill the remaining vertical space
        action_options_frame.pack(fill=tk.BOTH, expand=True)

    def _create_results_summary(self, parent):
        summary_outer_frame = tk.Frame(parent, bg=Theme.SECONDARY, relief="flat", bd=1)
        summary_outer_frame.pack(fill=tk.X)

        # Title
        title_frame = tk.Frame(summary_outer_frame, bg=Theme.TERTIARY, height=35)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="📊 Results Summary", font=("Segoe UI", 12, "bold"),
                fg=Theme.TEXT, bg=Theme.TERTIARY).pack(side=tk.LEFT, padx=15, pady=8)

        # Content area for summary - now a scrollable canvas
        content_frame = tk.Frame(summary_outer_frame, bg=Theme.SECONDARY)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15) # Pad content within outer frame

        self.summary_canvas = tk.Canvas(content_frame, bg=Theme.SECONDARY, highlightthickness=0)
        self.summary_scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.summary_canvas.yview)
        self.summary_frame_inner = tk.Frame(self.summary_canvas, bg=Theme.SECONDARY) # This holds the actual summary text

        self.summary_frame_inner.bind(
            "<Configure>",
            lambda e: self.summary_canvas.configure(
                scrollregion=self.summary_canvas.bbox("all")
            )
        )

        # Initial width for the window inside the canvas. It will be updated by _on_modal_resize.
        # Use a placeholder width for now, it will be corrected once modal is visible.
        self.summary_canvas.create_window((0, 0), window=self.summary_frame_inner, anchor="nw", width=1)
        self.summary_canvas.configure(yscrollcommand=self.summary_scrollbar.set)

        self.summary_scrollbar.pack(side="right", fill="y")
        self.summary_canvas.pack(side="left", fill="both", expand=True)

        # Bind mouse wheel for scrolling on this specific canvas
        def _on_summary_mousewheel(event):
            # Normalizes delta for different OS. Windows: 120, Mac: variable.
            # Using int(-event.delta/120) makes it work like a standard scroll.
            if event.num == 4 or event.delta > 0: # Scroll up
                self.summary_canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0: # Scroll down
                self.summary_canvas.yview_scroll(1, "units")
        self.summary_canvas.bind("<MouseWheel>", _on_summary_mousewheel) # Windows/macOS
        self.summary_canvas.bind("<Button-4>", _on_summary_mousewheel)  # Linux scroll up
        self.summary_canvas.bind("<Button-5>", _on_summary_mousewheel)  # Linux scroll down

        # File info (packed into the inner scrollable frame)
        file_frame = tk.Frame(self.summary_frame_inner, bg=Theme.SECONDARY)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(file_frame, text="📁 File:", font=("Segoe UI", 10, "bold"),
                fg=Theme.TEXT, bg=Theme.SECONDARY).pack(side=tk.LEFT)

        filename = Path(self.output_file).name if self.output_file else "Unknown"
        tk.Label(file_frame, text=filename, font=("Segoe UI", 10),
                fg=Theme.CYAN, bg=Theme.SECONDARY).pack(side=tk.LEFT, padx=(5, 0))

        # Analysis summary text (packed into the inner scrollable frame)
        summary = self.analysis_summary
        summary_text = f"""• {summary['total_proteins']} proteins processed
• UniProt data: {summary['uniprot_complete']}/{summary['total_proteins']} ({summary['uniprot_percent']:.1f}%)"""

        if summary.get('protparam_complete', 0) > 0:
            summary_text += f"\n• ProtParam data: {summary['protparam_complete']}/{summary['total_proteins']} ({summary['protparam_percent']:.1f}%)"

        if summary.get('blast_complete', 0) > 0:
            summary_text += f"\n• BLAST data: {summary['blast_complete']}/{summary['total_proteins']} ({summary['blast_percent']:.1f}%)"

        if summary.get('pdb_complete', 0) > 0:
            summary_text += f"\n• PDB structures: {summary['pdb_complete']}/{summary['total_proteins']} ({summary['pdb_percent']:.1f}%)"

        # *** IMPORTANT FIX: REMOVED DUMMY LINES ***
        # The dummy lines were here for testing and should NOT be in the final version.
        # for i in range(5):
        #     summary_text += f"\n• Additional line {i+1} to demonstrate scrolling capability and longer summaries."
        #     summary_text += f"\n• This ensures the summary section is the only one that scrolls, not the action options."
        # *** END FIX ***


        tk.Label(self.summary_frame_inner, text=summary_text, font=("Segoe UI", 10),
                fg=Theme.TEXT, bg=Theme.SECONDARY, justify=tk.LEFT).pack(anchor=tk.W)

        # Bind modal configure event to update canvas internal window width
        self.modal.bind("<Configure>", self._on_modal_resize)

        return summary_outer_frame

    def _on_modal_resize(self, event):
        """Adjusts the internal frame width inside the canvas when the modal resizes."""
        if self.summary_canvas and self.summary_frame_inner:
            # Calculate the effective width for the inner frame
            canvas_width = self.summary_canvas.winfo_width()
            scrollbar_width = self.summary_scrollbar.winfo_width() if self.summary_scrollbar.winfo_ismapped() else 0
            new_inner_width = canvas_width - scrollbar_width - 5 # A small buffer

            # Update the width of the window created inside the canvas
            # Ensure the item exists before configuring
            if self.summary_canvas.find_all():
                self.summary_canvas.itemconfig(self.summary_canvas.find_all()[0], width=new_inner_width)
            else:
                # If for some reason create_window hasn't run yet, try again.
                # This should ideally not happen if show() calls update_idletasks.
                self.summary_canvas.create_window((0, 0), window=self.summary_frame_inner, anchor="nw", width=new_inner_width)

            # Update the scroll region to reflect potential new content size
            self.summary_canvas.configure(scrollregion=self.summary_canvas.bbox("all"))


    def _create_action_options(self, parent):
        options_frame = tk.Frame(parent, bg=Theme.SECONDARY, relief="solid", bd=1)
        # This frame should fill the remaining vertical space after the summary and before bottom buttons
        options_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_frame = tk.Frame(options_frame, bg=Theme.CYAN, height=40)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="🚀 What would you like to do next?", font=("Segoe UI", 12, "bold"),
                fg="white", bg=Theme.CYAN).pack(side=tk.LEFT, padx=15, pady=10)

        # Options grid container
        options_content = tk.Frame(options_frame, bg=Theme.SECONDARY)
        options_content.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Configure grid weights - ensuring cells expand
        options_content.grid_rowconfigure(0, weight=1)
        options_content.grid_rowconfigure(1, weight=1)
        options_content.grid_columnconfigure(0, weight=1)
        options_content.grid_columnconfigure(1, weight=1)

        # Option 1: Similarity Analysis
        self._create_action_option(options_content,
                                "🔬 Similarity Analysis",
                                "Compare proteins to find similar ones based on properties",
                                self._similarity_analysis, 0, 0)

        # Option 2: View Results
        self._create_action_option(options_content,
                                "📊 View Results",
                                "Open and explore your analysis results",
                                self._view_results, 0, 1)

        # Option 3: Data Viewer
        self._create_action_option(options_content,
                                "📈 Data Viewer",
                                "Browse all sheets and data in detail",
                                self._data_viewer, 1, 0)

        # Option 4: Export Data
        self._create_action_option(options_content,
                                "💾 Export Options",
                                "Export data to different formats",
                                self._export_options, 1, 1)
        return options_frame

    def _create_action_option(self, parent, title, description, command, row, col):
        option_frame = tk.Frame(parent, bg=Theme.TERTIARY, relief="solid", bd=1)
        option_frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew") # Added sticky

        # Ensure elements within option_frame also expand to fill available space
        option_frame.grid_rowconfigure(0, weight=1) # For text content
        option_frame.grid_rowconfigure(1, weight=0) # For button (fixed size)
        option_frame.grid_columnconfigure(0, weight=1)

        # Inner frame for title and description to allow flexible wrapping
        text_content_frame = tk.Frame(option_frame, bg=Theme.TERTIARY)
        text_content_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 5))

        # Title
        tk.Label(text_content_frame, text=title, font=("Segoe UI", 11, "bold"),
                fg=Theme.TEXT, bg=Theme.TERTIARY).pack(pady=(0, 5))

        # Description - use wraplength and pack with fill/expand to allow resizing
        # The wraplength should be dynamic, but fixed for now.
        # Adjusted wraplength to provide enough room
        tk.Label(text_content_frame, text=description, font=("Segoe UI", 9),
                fg=Theme.TEXT_MUTED, bg=Theme.TERTIARY, wraplength=180,
                justify=tk.CENTER).pack(fill=tk.BOTH, expand=True, padx=10)


        # Button - Ensure button packs correctly within its frame at the bottom
        btn = ModernButton(option_frame, "Select", command, "secondary")
        btn.pack(pady=(0, 10), padx=20) # Packed centrally at the bottom of the option_frame

        # Additional click binding for reliability
        def on_click(event=None):
            try:
                command()
            except Exception as e:
                print(f"Button click error: {e}")

        btn.bind('<Button-1>', on_click)

    def _create_bottom_buttons(self):
        button_area = tk.Frame(self.modal, bg=Theme.BG, height=80)
        button_area.pack(fill=tk.X, side=tk.BOTTOM, expand=False) # Pack at bottom, don't expand vertically

        # Ensure the content within button_area is centered
        button_container = tk.Frame(button_area, bg=Theme.BG)
        button_container.pack(expand=True, fill=tk.BOTH, padx=30, pady=20) # Allow inner container to expand within button_area

        button_frame = tk.Frame(button_container, bg=Theme.BG)
        button_frame.pack(anchor=tk.CENTER, expand=False) # Do not expand this frame, just center its content

        ModernButton(button_frame, "📁 Open Results File", self._open_results, "secondary").pack(side=tk.LEFT, padx=(0, 15))
        ModernButton(button_frame, "🔄 Return to Main Menu", self._new_analysis, "warning").pack(side=tk.LEFT, padx=(0, 15))
        ModernButton(button_frame, "✅ Finish", self._finish, "success").pack(side=tk.LEFT)

    def _similarity_analysis(self):
        try:
            self.result = "similarity"
            if self.modal:
                self.modal.destroy()
        except Exception as e:
            print(f"Similarity analysis error: {e}")

    def _view_results(self):
        try:
            self.result = "view_results"
            if self.modal:
                self.modal.destroy()
        except Exception as e:
            print(f"View results error: {e}")

    def _data_viewer(self):
        try:
            self.result = "data_viewer"
            if self.modal:
                self.modal.destroy()
        except Exception as e:
            print(f"Data viewer error: {e}")

    def _export_options(self):
        try:
            self.result = "export_options"
            if self.modal:
                self.modal.destroy()
        except Exception as e:
            print(f"Export options error: {e}")

    def _open_results(self):
        def open_file_async():
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(str(self.output_file))
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.run(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', str(self.output_file)])
            except Exception as e:
                if self.modal:
                    self.modal.after(0, lambda: messagebox.showerror("Error", f"{ERROR_MESSAGES['FILE_OPEN_ERROR']} {e}"))

        threading.Thread(target=open_file_async, daemon=True).start()

    def _new_analysis(self):
        try:
            self.result = "new_analysis"
            if self.modal:
                self.modal.destroy()
        except Exception as e:
            print(f"New analysis error: {e}")

    def _finish(self):
        try:
            self.result = "finish"
            if self.modal:
                self.modal.destroy()
        except Exception as e:
            print(f"Finish error: {e}")


class OptionsModal:
    """Analysis options configuration dialog"""

    def __init__(self, parent, current_options=None):
        self.parent = parent
        self.result = None
        self.current_options = current_options or {}
    
        # Get current options, ensuring gene_ids state is properly synced
        opts = self.current_options
        self.protparam_var = tk.BooleanVar(value=opts.get('protparam', True))
        self.blast_var = tk.BooleanVar(value=opts.get('blast', False))
        self.amino_acid_var = tk.BooleanVar(value=opts.get('amino_acid', False))
        self.pdb_var = tk.BooleanVar(value=opts.get('pdb_search', False))
        self.safe_mode_var = tk.BooleanVar(value=opts.get('safe_mode', True))
        # This will now properly reflect the current GUI state
        self.gene_id_var = tk.BooleanVar(value=opts.get('use_gene_ids', False))

    def show(self):
        self.modal = tk.Toplevel(self.parent)
        self.modal.title("Analysis Options")
        self.modal.geometry("550x650")
        self.modal.configure(bg=Theme.BG)
        self.modal.resizable(False, False)
        self.modal.transient(self.parent)
        self.modal.grab_set()

        opts = self.current_options
        self.protparam_var = tk.BooleanVar(value=opts.get('protparam', True))
        self.blast_var = tk.BooleanVar(value=opts.get('blast', False))
        self.amino_acid_var = tk.BooleanVar(value=opts.get('amino_acid', False))
        self.pdb_var = tk.BooleanVar(value=opts.get('pdb_search', False))
        self.safe_mode_var = tk.BooleanVar(value=opts.get('safe_mode', True))

        # Center on parent
        self.modal.update_idletasks()
        x = (self.parent.winfo_x() + self.parent.winfo_width() // 2) - (self.modal.winfo_width() // 2)
        y = (self.parent.winfo_y() + self.parent.winfo_height() // 2) - (self.modal.winfo_height() // 2)
        self.modal.geometry(f"+{x}+{y}")

        self._create_content()
        self.parent.wait_window(self.modal)
        return self.result

    def _create_content(self):
        # Header
        header = tk.Frame(self.modal, bg=Theme.CYAN, height=50)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text="Analysis Options", font=("Segoe UI", 16, "bold"),
                fg="white", bg=Theme.CYAN).pack(expand=True)

        # Content
        content = tk.Frame(self.modal, bg=Theme.BG)
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # Basic section
        basic = self._create_section(content, "Basic Analysis")
        self._create_option(basic, "UniProt Data", "Core protein information", None, enabled=False)
        self._create_option(basic, "ProtParam Analysis", "Molecular properties", self.protparam_var)
        self._create_option(basic, "  └─ Amino Acid Details", "Detailed composition", self.amino_acid_var)

        # Advanced section
        advanced = self._create_section(content, "Advanced Analysis")
        self._create_option(advanced, "BLAST Search", "Slow: ~1-2 min/protein", self.blast_var, warning=True)
        self._create_option(advanced, "PDB Structures", "3D structure data", self.pdb_var)

        # Settings section - modify this part
        settings = self._create_section(content, "Settings")
        self._create_option(settings, "Safe Mode", "Preserve existing data", self.safe_mode_var)
        self._create_option(settings, "Use Gene IDs", "Input uses gene names instead of UniProt IDs", self.gene_id_var)

        # Buttons
        btn_frame = tk.Frame(self.modal, bg=Theme.BG)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 15))

        ModernButton(btn_frame, "Cancel", self._cancel, "secondary").pack(side=tk.RIGHT, padx=(10, 0))
        ModernButton(btn_frame, "Apply", self._apply, "success").pack(side=tk.RIGHT)

    def _create_section(self, parent, title):
        section = tk.Frame(parent, bg=Theme.SECONDARY)
        section.pack(fill=tk.X, pady=(0, 10))

        title_frame = tk.Frame(section, bg=Theme.TERTIARY, height=30)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text=title, font=("Segoe UI", 11, "bold"),
                fg=Theme.TEXT, bg=Theme.TERTIARY).pack(side=tk.LEFT, padx=10, pady=5)

        content = tk.Frame(section, bg=Theme.SECONDARY)
        content.pack(fill=tk.X, padx=10, pady=8)
        return content

    def _create_option(self, parent, text, desc, variable, enabled=True, warning=False):
        frame = tk.Frame(parent, bg=Theme.SECONDARY)
        frame.pack(fill=tk.X, pady=2)

        if variable:
            style = ttk.Style()
            style.configure('Option.TCheckbutton', background=Theme.SECONDARY,
                          foreground=Theme.TEXT, focuscolor='none')

            cb = ttk.Checkbutton(frame, text=text, variable=variable, style='Option.TCheckbutton',
                               state="normal" if enabled else "disabled")
            cb.pack(side=tk.LEFT, anchor="w")
        else:
            tk.Label(frame, text=text, font=("Segoe UI", 10, "bold"),
                    fg=Theme.GREEN, bg=Theme.SECONDARY).pack(side=tk.LEFT, anchor="w")

        desc_color = Theme.RED if warning else Theme.TEXT_MUTED
        tk.Label(frame, text=desc, font=("Segoe UI", 9), fg=desc_color, bg=Theme.SECONDARY,
                wraplength=300, justify=tk.LEFT).pack(side=tk.LEFT, padx=(10, 0), anchor="w")

    def _apply(self):
        self.result = {
            'uniprot': True,
            'protparam': self.protparam_var.get(),
            'blast': self.blast_var.get(),
            'amino_acid': self.amino_acid_var.get() and self.protparam_var.get(),
            'pdb_search': self.pdb_var.get(),
            'safe_mode': self.safe_mode_var.get(),
            'use_gene_ids': self.gene_id_var.get()  # NEW
        }
        self.modal.destroy()

    def _cancel(self):
        self.result = None
        self.modal.destroy()


class ProtMergeGUI:
    """Main ProtMerge GUI with integrated features"""

    def __init__(self, protmerge_app):
        self.app = protmerge_app
        self.logger = logging.getLogger(__name__)

        # State
        self.input_file = None
        self.uniprot_column = None
        self.file_selected = False
        self.column_selected = False
        self.options_configured = False

        # Similarity state
        self.similarity_file = None
        self.similarity_file_selected = False

        # Current options
        self.current_options = {
            'uniprot': True, 'protparam': True, 'blast': False,
            'amino_acid': False, 'pdb_search': False, 'safe_mode': True
        }
    
        # NEW: Add gene ID toggle variable
        self.use_gene_ids = None  # Will be initialized in run()

        # GUI elements
        self.root = None
        self.sheet_var = None
        self.progress_bar = None
        self.progress_frame = None

        # Tab buttons
        self.new_analysis_btn = None
        self.similarity_btn = None
        self.results_btn = None

        # Similarity components
        self.similarity_file_label = None
        self.similarity_info_label = None
        self.similarity_status = None
        self.launch_similarity_btn = None

        # Current mode
        self.current_mode = "new_analysis"

        #Button Fix
        self.completion_dialog = None
        self.analysis_in_progress = False

    def run(self):
        """Initialize and run GUI"""
        try:
            self._setup_window()
            self.sheet_var = tk.StringVar()
            self.use_gene_ids = tk.BooleanVar(value=False)  # Initialize here
            self._create_widgets()
            self._center_window()
            self.logger.info("GUI initialized successfully")
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f"GUI failed to start: {e}")
            raise

    def _setup_window(self):
        self.root = tk.Tk()
        self.root.title("ProtMerge v1.2.0 - Protein Analysis Suite")
        self.root.geometry("850x750")
        self.root.configure(bg=Theme.BG)
        self.root.resizable(False, False)

    def _create_widgets(self):
        """Create main interface with integrated features"""
        main = tk.Frame(self.root, bg=Theme.BG)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # Header
        self._create_header(main)

        # Mode switcher
        self._create_mode_switcher(main)

        # Content area that changes based on mode
        self.content_frame = tk.Frame(main, bg=Theme.BG)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        # Create initial content (new analysis mode)
        self._create_new_analysis_content()

    def _create_header(self, parent):
        """Create application header"""
        header = tk.Frame(parent, bg=Theme.BG)
        header.pack(fill=tk.X, pady=(0, 20))

        tk.Label(header, text="ProtMerge", font=("Segoe UI", 20, "bold"),
                fg=Theme.TEXT, bg=Theme.BG).pack()
        tk.Label(header, text="Protein Analysis Suite v1.2.0", font=("Segoe UI", 10),
                fg=Theme.TEXT_MUTED, bg=Theme.BG).pack(pady=(5, 0))

        tk.Frame(header, height=2, bg=Theme.CYAN).pack(fill=tk.X, pady=(10, 0))

    def _create_mode_switcher(self, parent):
        """Create mode switcher tabs"""
        switcher_frame = tk.Frame(parent, bg=Theme.BG)
        switcher_frame.pack(fill=tk.X, pady=(0, 15))

        # Create tab-like buttons
        tab_frame = tk.Frame(switcher_frame, bg=Theme.BG)
        tab_frame.pack()

        # New Analysis tab - start all as ghost, will be updated by _update_tab_styles
        self.new_analysis_btn = ModernButton(tab_frame, "🧬 New Analysis",
                                        self._switch_to_new_analysis,
                                        "ghost")
        self.new_analysis_btn.pack(side=tk.LEFT, padx=(0, 5))

        # Similarity Analyzer tab
        self.similarity_btn = ModernButton(tab_frame, "🔬 Similarity Analyzer",
                                        self._switch_to_similarity,
                                        "ghost")
        self.similarity_btn.pack(side=tk.LEFT, padx=(0, 5))

        # Results Manager tab
        self.results_btn = ModernButton(tab_frame, "📊 Results Manager",
                                    self._switch_to_results,
                                    "ghost")
        self.results_btn.pack(side=tk.LEFT)

        # Update initial active state
        self._update_tab_styles()

    def _switch_to_new_analysis(self):
        """Switch to new analysis mode"""
        self.current_mode = "new_analysis"
        self._update_tab_styles()
        self._clear_content()
        self._create_new_analysis_content()

    def _switch_to_similarity(self):
        """Switch to similarity analyzer mode"""
        self.current_mode = "similarity_mode"
        self._update_tab_styles()
        self._clear_content()
        self._create_similarity_content()

    def _switch_to_results(self):
        """Switch to results manager mode"""
        self.current_mode = "results_mode"
        self._update_tab_styles()
        self._clear_content()
        self._create_results_content()

    def _update_tab_styles(self):
        """Update tab button styles based on current mode"""
        # Ensure that ModernButton's set_nav_active method is used consistently.
        # This method handles setting the correct background and foreground colors.
        if self.new_analysis_btn:
            self.new_analysis_btn.set_nav_active(self.current_mode == "new_analysis")

        if self.similarity_btn:
            self.similarity_btn.set_nav_active(self.current_mode == "similarity_mode")

        if self.results_btn:
            self.results_btn.set_nav_active(self.current_mode == "results_mode")

    def _clear_content(self):
        """Clear current content"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def _create_new_analysis_content(self):
        """Create new analysis interface"""
        # File section (unchanged)
        file_section = self._create_section(self.content_frame, "📁 Input File")
        self.file_label = tk.Label(file_section, text="No file selected",
                                fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY, anchor="w")
        self.file_label.pack(fill=tk.X, padx=10, pady=(0, 10))

        file_controls = tk.Frame(file_section, bg=Theme.SECONDARY)
        file_controls.pack(fill=tk.X, padx=10, pady=(0, 10))

        ModernButton(file_controls, "Browse", self._browse_file).pack(side=tk.LEFT)
        tk.Label(file_controls, text="Sheet:", fg=Theme.TEXT, bg=Theme.SECONDARY).pack(side=tk.LEFT, padx=(15, 5))

        self.sheet_combo = ttk.Combobox(file_controls, textvariable=self.sheet_var,
                                    state="disabled", width=15)
        self.sheet_combo.pack(side=tk.LEFT)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._load_columns)

        # Column section with gene ID toggle - FIXED VERSION
        col_section = self._create_section(self.content_frame, "🎯 UniProt Column")
    
        # Column header with toggle button
        header_frame = tk.Frame(col_section, bg=Theme.SECONDARY)
        header_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
    
        # Column selection label
        self.column_label = tk.Label(
            header_frame, 
            text="Select column with UniProt IDs:",
            fg=Theme.TEXT, 
            bg=Theme.SECONDARY, 
            anchor="w"
        )
        self.column_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
        # Toggle button for gene IDs
        self.gene_toggle = ModernButton(
            header_frame, 
            "Gene IDs", 
            self._toggle_gene_mode, 
            "ghost", 
            "small"
        )
        self.gene_toggle.pack(side=tk.RIGHT)

        # Column listbox (unchanged)
        list_frame = tk.Frame(col_section, bg=Theme.SECONDARY)
        list_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.column_listbox = tk.Listbox(list_frame, height=4, bg=Theme.TERTIARY, fg=Theme.TEXT,
                                        selectbackground=Theme.CYAN, selectforeground="white",
                                        relief="flat", borderwidth=0)
        scrollbar = tk.Scrollbar(list_frame, command=self.column_listbox.yview)
        self.column_listbox.configure(yscrollcommand=scrollbar.set)
        self.column_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.column_listbox.bind("<<ListboxSelect>>", self._column_selected)

        # Options section (unchanged)
        opt_section = self._create_section(self.content_frame, "⚙️ Analysis Options")
        opt_frame = tk.Frame(opt_section, bg=Theme.SECONDARY)
        opt_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.options_summary = tk.Label(opt_frame, text="Default: UniProt + ProtParam",
                                    fg=Theme.TEXT, bg=Theme.SECONDARY, anchor="w")
        self.options_summary.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ModernButton(opt_frame, "Configure", self._show_options, "secondary").pack(side=tk.RIGHT)

        # Progress section (initially hidden)
        self.progress_frame = self._create_section(self.content_frame, "📊 Progress")
        self.progress_frame.pack_forget()  # Hide initially

        self.progress_bar = ProgressBar(self.progress_frame)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(0, 5))

        self.progress_text = tk.Label(self.progress_frame, text="", font=("Segoe UI", 10, "bold"),
                                    fg=Theme.TEXT, bg=Theme.SECONDARY, anchor="w")
        self.progress_text.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Bottom controls for new analysis
        self._create_new_analysis_controls()

        # Update options summary
        self._update_options_summary()

    def _create_similarity_content(self):
        """Create similarity analyzer interface"""
        # Info section
        info_section = self._create_section(self.content_frame, "🔬 Similarity Analyzer")

        info_frame = tk.Frame(info_section, bg=Theme.SECONDARY)
        info_frame.pack(fill=tk.X, padx=15, pady=15)

        tk.Label(info_frame, text="Analyze protein similarities using existing ProtMerge results",
                font=("Segoe UI", 12, "bold"), fg=Theme.TEXT, bg=Theme.SECONDARY).pack(anchor=tk.W, pady=(0, 10))

        features_text = """Features:
• Compare proteins based on sequence, molecular properties, and function
• Multiple analysis presets (Basic, Sequence-focused, Biochemical, Functional)
• Interactive results viewer with export capabilities
• Real-time similarity scoring and ranking"""

        tk.Label(info_frame, text=features_text, font=("Segoe UI", 10),
                fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY, justify=tk.LEFT).pack(anchor=tk.W)

        # File selection section
        file_section = self._create_section(self.content_frame, "📁 Select Results File")

        self.similarity_file_label = tk.Label(file_section, text="No ProtMerge results file selected",
                                             fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY, anchor="w")
        self.similarity_file_label.pack(fill=tk.X, padx=10, pady=(0, 10))

        file_controls = tk.Frame(file_section, bg=Theme.SECONDARY)
        file_controls.pack(fill=tk.X, padx=10, pady=(0, 10))

        ModernButton(file_controls, "Browse Results File", self._browse_similarity_file).pack(side=tk.LEFT)

        self.similarity_info_label = tk.Label(file_controls, text="",
                                             fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY)
        self.similarity_info_label.pack(side=tk.LEFT, padx=(15, 0))

        # Quick actions
        actions_section = self._create_section(self.content_frame, "⚡ Quick Actions")

        actions_frame = tk.Frame(actions_section, bg=Theme.SECONDARY)
        actions_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Recent files (placeholder)
        tk.Label(actions_frame, text="Recent ProtMerge files will appear here",
                font=("Segoe UI", 10), fg=Theme.TEXT_MUTED, bg=Theme.SECONDARY).pack(pady=10)

        # Bottom controls for similarity
        self._create_similarity_controls()

    def _create_results_content(self):
        """Create results manager interface"""
        # Info section
        info_section = self._create_section(self.content_frame, "📊 Results Manager")

        info_frame = tk.Frame(info_section, bg=Theme.SECONDARY)
        info_frame.pack(fill=tk.X, padx=15, pady=15)

        tk.Label(info_frame, text="Manage and explore your ProtMerge analysis results",
                font=("Segoe UI", 12, "bold"), fg=Theme.TEXT, bg=Theme.SECONDARY).pack(anchor=tk.W, pady=(0, 10))

        # Tools grid
        tools_section = self._create_section(self.content_frame, "🛠️ Available Tools")

        tools_frame = tk.Frame(tools_section, bg=Theme.SECONDARY)
        tools_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Configure grid
        for i in range(3):
            tools_frame.grid_columnconfigure(i, weight=1)
        for i in range(2):
            tools_frame.grid_rowconfigure(i, weight=1)

        # Tool options
        tools = [
            ("📊 Data Viewer", "Browse and explore analysis results", self._launch_data_viewer, 0, 0),
            ("📈 Statistics", "Generate summary statistics", self._show_statistics, 0, 1),
            ("🔍 Search Proteins", "Find specific proteins", self._search_proteins, 0, 2),
            ("💾 Export Data", "Convert to different formats", self._export_data, 1, 0),
            ("🔄 Merge Files", "Combine multiple results", self._merge_files, 1, 1),
            ("📋 Compare Results", "Compare different analyses", self._compare_results, 1, 2)
        ]

        for title, desc, command, row, col in tools:
            self._create_tool_option(tools_frame, title, desc, command, row, col)

    def _create_tool_option(self, parent, title, description, command, row, col):
        """Create a tool option in the grid"""
        tool_frame = tk.Frame(parent, bg=Theme.TERTIARY, relief="solid", bd=1)
        tool_frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

        # Title
        tk.Label(tool_frame, text=title, font=("Segoe UI", 11, "bold"),
                fg=Theme.TEXT, bg=Theme.TERTIARY).pack(pady=(15, 5))

        # Description
        tk.Label(tool_frame, text=description, font=("Segoe UI", 9),
                fg=Theme.TEXT_MUTED, bg=Theme.TERTIARY, wraplength=150,
                justify=tk.CENTER).pack(pady=(0, 10), padx=10)

        # Button
        ModernButton(tool_frame, "Launch", command, "secondary").pack(pady=(0, 15))

    def _create_section(self, parent, title):
        """Create a styled section"""
        section = tk.Frame(parent, bg=Theme.SECONDARY)
        section.pack(fill=tk.X, pady=(0, 12))

        title_frame = tk.Frame(section, bg=Theme.TERTIARY, height=30)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text=title, font=("Segoe UI", 11, "bold"),
                fg=Theme.TEXT, bg=Theme.TERTIARY).pack(side=tk.LEFT, padx=10, pady=6)

        return section

    def _create_new_analysis_controls(self):
        """Create controls for new analysis mode"""
        tk.Frame(self.content_frame, bg=Theme.BG, height=10).pack(fill=tk.X)

        controls = tk.Frame(self.content_frame, bg=Theme.BG)
        controls.pack(fill=tk.X, pady=(15, 0))

        self.status_label = tk.Label(controls, text="Select file and column to continue",
                                    fg=Theme.TEXT_MUTED, bg=Theme.BG, anchor="w")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.start_btn = ModernButton(controls, "START ANALYSIS", self._start_analysis, "success")
        self.start_btn.pack(side=tk.RIGHT)

        self.cancel_btn = ModernButton(controls, "Cancel", self._cancel, "danger")
        self.cancel_btn.pack(side=tk.RIGHT, padx=(0, 10))

        self._update_start_button()

    def _create_similarity_controls(self):
        """Create controls for similarity mode"""
        controls = tk.Frame(self.content_frame, bg=Theme.BG)
        controls.pack(fill=tk.X, pady=(20, 0))

        self.similarity_status = tk.Label(controls, text="Select a ProtMerge results file to begin",
                                         fg=Theme.TEXT_MUTED, bg=Theme.BG, anchor="w")
        self.similarity_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.launch_similarity_btn = ModernButton(controls, "🔬 Launch Similarity Analyzer",
                                                 self._launch_similarity_analyzer, "primary")
        self.launch_similarity_btn.pack(side=tk.RIGHT)

        # Initially disable the button
        self.launch_similarity_btn.config(state="disabled", bg="#4a4a4a")
        self._update_similarity_button_state()
        
    def _toggle_gene_mode(self):
        """Toggle between UniProt ID and Gene ID input modes"""
        current_state = self.use_gene_ids.get()
        self.use_gene_ids.set(not current_state)

        self.current_options['use_gene_ids'] = self.use_gene_ids.get()

        self._update_gene_toggle_appearance()

        self._update_options_summary()

    # =============================================================================
    # FILE HANDLING METHODS
    # =============================================================================

    def _browse_file(self):
        """Browse for input file (new analysis)"""
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=SUPPORTED_EXCEL_FORMATS)
        if file_path:
            self.input_file = Path(file_path)
            self.file_label.config(text=f"Selected: {self.input_file.name}", fg=Theme.GREEN)

            try:
                xl_file = pd.ExcelFile(self.input_file)
                self.sheet_combo['values'] = xl_file.sheet_names
                self.sheet_combo['state'] = 'readonly'
                if xl_file.sheet_names:
                    self.sheet_combo.set(xl_file.sheet_names[0])
                    self._load_columns()
                self.file_selected = True
                self._update_status("File loaded. Select UniProt column.", Theme.CYAN)
            except Exception as e:
                messagebox.showerror("Error", f"{ERROR_MESSAGES['FILE_LOAD_ERROR']} {e}")
                self.file_selected = False

            self._update_start_button()

    def _browse_similarity_file(self):
        """Browse for similarity analysis file"""
        file_path = filedialog.askopenfilename(
            title="Select ProtMerge Results File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            if self._validate_protmerge_file(file_path):
                self.similarity_file = Path(file_path)
                self.similarity_file_selected = True

                # Update UI
                self.similarity_file_label.config(text=f"Selected: {self.similarity_file.name}", fg=Theme.GREEN)

                # Get file info
                try:
                    df = pd.read_excel(file_path, sheet_name='ProtMerge_Results', nrows=5) # Read small sample
                    # Re-read full dataframe to get accurate count if needed, or rely on info from validation
                    full_df = pd.read_excel(file_path, sheet_name='ProtMerge_Results')
                    protein_count = len(full_df)
                    self.similarity_info_label.config(text=f"{protein_count} proteins", fg=Theme.CYAN)
                    self.similarity_status.config(text="Ready to launch similarity analysis", fg=Theme.GREEN)

                    # Enable the button
                    self._update_similarity_button_state()

                except Exception as e:
                    self.similarity_info_label.config(text="Error reading file", fg=Theme.RED)
                    self.similarity_status.config(text="File error - please select a different file", fg=Theme.RED)
                    self.similarity_file_selected = False
                    self._update_similarity_button_state()

    def _validate_protmerge_file(self, file_path):
        """Validate ProtMerge results file"""
        try:
            xl_file = pd.ExcelFile(file_path)

            if 'ProtMerge_Results' not in xl_file.sheet_names:
                messagebox.showerror(
                    "Invalid File",
                    f"{ERROR_MESSAGES['INVALID_PROTMERGE_FILE']}\n\n"
                    f"{ERROR_MESSAGES['SHEET_NOT_FOUND']}"
                )
                return False

            df = pd.read_excel(file_path, sheet_name='ProtMerge_Results', nrows=5)
            if 'UniProt ID' not in df.columns:
                messagebox.showerror(
                    "Invalid Format",
                    f"{ERROR_MESSAGES['REQUIRED_COLUMN_MISSING']}"
                )
                return False

            # Check total rows after confirming sheet and column exist
            total_rows = len(pd.read_excel(file_path, sheet_name='ProtMerge_Results'))
            if total_rows < 2:
                messagebox.showwarning(
                    "Insufficient Data",
                    f"{ERROR_MESSAGES['INSUFFICIENT_DATA']}\n"
                    f"This file contains {total_rows} proteins."
                )
                return False

            return True

        except Exception as e:
            messagebox.showerror("File Error", f"{ERROR_MESSAGES['FILE_LOAD_ERROR']}\n{e}")
            return False

    def _load_columns(self, event=None):
        """Load columns for new analysis"""
        if not self.input_file or not self.sheet_var.get():
            return

        try:
            df = pd.read_excel(self.input_file, sheet_name=self.sheet_var.get(), nrows=5)
            self.column_listbox.delete(0, tk.END)

            for i, col in enumerate(df.columns):
                sample = df[col].dropna().head(2).tolist()
                sample_str = ", ".join(str(x)[:15] for x in sample) if sample else "Empty"
                if len(sample_str) > 40:
                    sample_str = sample_str[:37] + "..."

                self.column_listbox.insert(tk.END, f"{chr(65+i)} | {col} | {sample_str}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load columns: {e}")

    def _column_selected(self, event):
        """Handle column selection"""
        selection = self.column_listbox.curselection()
        if selection:
            self.uniprot_column = selection[0]
            self.column_selected = True
            col_name = self.column_listbox.get(selection[0]).split('|')[1].strip()
            self._update_status(f"Selected: {col_name}", Theme.GREEN)
            self._update_start_button()

    # =============================================================================
    # ANALYSIS METHODS
    # =============================================================================

    def _start_analysis(self):
        """Start new protein analysis"""
        if self.current_options.get('blast') and not messagebox.askyesno(
            "BLAST Warning", "BLAST is slow (~1-2 min/protein). Continue?"):
            return

        self.logger.info("Starting analysis thread...")
        self._show_progress()

        analysis_thread = threading.Thread(target=self._run_analysis, daemon=True)
        analysis_thread.start()
        self.logger.info("Analysis thread started")

    def _show_progress(self):
        """Show progress section"""
        self.analysis_in_progress = True

        # Show progress
        if hasattr(self, 'progress_frame'):
            self.progress_frame.pack(fill=tk.X, pady=(0, 12))

        if hasattr(self, 'start_btn'):
            self.start_btn.config(state="disabled", text="ANALYZING...", bg="#4a4a4a")

        self._update_status("Analysis running...", Theme.CYAN)
        self.root.update_idletasks()

    def _run_analysis(self):
        """Run analysis in background thread"""
        try:
            def progress_callback(pct, main_text, detail_text=""):
                adjusted_progress = self._calculate_smooth_progress(pct, main_text)
                self.root.after(0, self._update_progress, adjusted_progress, main_text, detail_text)

            output_file = self.app.run_analysis(
                self.input_file, self.sheet_var.get(), self.uniprot_column,
                self.current_options, progress_callback
            )

            analysis_summary = self.app.get_analysis_summary()
            self.root.after(0, self._show_completion, output_file, analysis_summary)

        except Exception as e:
            self.root.after(0, self._show_error, str(e))

    def _calculate_smooth_progress(self, raw_progress, main_text):
        """Calculate smooth progress based on enabled analyses"""
        stage = self._identify_stage(main_text.lower())

        stage_weights = {'uniprot': 40}
        total_weight = 40

        if self.current_options.get('protparam', False):
            stage_weights['protparam'] = 25
            total_weight += 25

        if self.current_options.get('blast', False):
            stage_weights['blast'] = 20
            total_weight += 20

        if self.current_options.get('pdb_search', False):
            stage_weights['pdb'] = 10
            total_weight += 10

        # Normalize weights
        for s in stage_weights:
            stage_weights[s] = (stage_weights[s] / total_weight) * 100

        # Calculate cumulative progress
        cumulative = 0
        stage_order = ['uniprot', 'protparam', 'blast', 'pdb']

        for i, s in enumerate(stage_order):
            if s in stage_weights:
                if s == stage:
                    stage_progress = min(max(raw_progress, 0), 100) / 100
                    return min(cumulative + (stage_weights[s] * stage_progress), 99)
                elif stage_order.index(s) < stage_order.index(stage):
                    cumulative += stage_weights[s]

        return min(max(raw_progress, 0), 99)

    def _identify_stage(self, text):
        """Identify current analysis stage"""
        if 'uniprot' in text or 'fetching' in text:
            return 'uniprot'
        elif 'protparam' in text:
            return 'protparam'
        elif 'blast' in text:
            return 'blast'
        elif 'pdb' in text:
            return 'pdb'
        else:
            return 'uniprot'

    def _update_progress(self, pct, text, detail_text=""):
        """Update progress display"""
        try:
            self.progress_bar.set_progress(pct)
            self.progress_text.config(text=text)
            # The progress bar and label's update_idletasks methods are called within ProgressBar.set_progress
            self.root.update_idletasks() # Ensure main window updates
        except Exception as e:
            self.logger.error(f"Progress update error: {e}")

    def _show_completion(self, output_file, analysis_summary):
        """Show completion dialog with options"""
        # Hide progress first
        self._hide_progress()

        self.completion_dialog = CompletionDialog(self.root, output_file, analysis_summary)
        action = self.completion_dialog.show()

        # Clear completion dialog reference
        self.completion_dialog = None

        if action == "similarity":
            self._launch_similarity_on_file(output_file)
            self._reset_for_new_analysis() # Go back to new analysis screen after launching similarity
        elif action == "view_results":
            self._open_file(output_file)
            self._reset_for_new_analysis() # Go back to new analysis screen
        elif action == "data_viewer":
            self._show_data_viewer(output_file)
            # Data viewer handles returning focus, do not reset main GUI state here
        elif action == "export_options":
            self._show_export_options(output_file)
            self._reset_for_new_analysis() # Go back to new analysis screen
        elif action == "new_analysis":
            self._reset_for_new_analysis()
        elif action == "finish":
            self.root.quit()

    def _show_error(self, error):
        """Show error dialog"""
        messagebox.showerror("Error", f"{ERROR_MESSAGES['ANALYSIS_FAILED']} {error}")
        self._hide_progress() # Ensure progress is hidden on error
        self._reset_for_new_analysis() # Reset state
        # self.root.quit() # Removed immediate quit, let user decide

    def _hide_progress(self):
        """Hide progress section and reset interface"""
        if hasattr(self, 'progress_frame') and self.progress_frame:
            self.progress_frame.pack_forget()

        if hasattr(self, 'start_btn'):
            self.start_btn.config(
                state="normal",
                text="START ANALYSIS",
                bg=Theme.GREEN,
                activebackground="#45a049"
            )

        self.analysis_in_progress = False

    def _return_to_main_menu(self):
        """Return to main menu and reset state"""
        self._hide_progress()
        self._update_status("Analysis complete. Ready for new analysis.", Theme.GREEN)
        self._switch_to_new_analysis() # Explicitly switch to new analysis tab

    # =============================================================================
    # SIMILARITY ANALYSIS METHODS
    # =============================================================================

    def _update_similarity_button_state(self):
        """Update similarity button state"""
        if hasattr(self, 'launch_similarity_btn') and self.launch_similarity_btn:
            if self.similarity_file_selected:
                self.launch_similarity_btn.config(
                    state="normal",
                    bg=Theme.CYAN,
                    activebackground=Theme.CYAN_HOVER
                )
            else:
                self.launch_similarity_btn.config(
                    state="disabled",
                    bg="#4a4a4a", # Muted disabled color
                    activebackground="#4a4a4a" # Disabled active background
                )

    def _launch_similarity_analyzer(self):
        """Launch similarity analyzer on selected file"""
        if hasattr(self, 'similarity_file') and self.similarity_file:
            self._launch_similarity_on_file(self.similarity_file)
        else:
            messagebox.showwarning("No File", "Please select a ProtMerge results file first.")

    def _launch_similarity_on_file(self, file_path):
        """Launch similarity analysis on specific file"""
        try:
            # The import is now handled at the top of the file,
            # using mock if the actual module is not found.
            # from similarity_gui import launch_similarity_analysis # Moved to top

            file_name = Path(file_path).name
            total_proteins = len(pd.read_excel(file_path, sheet_name='ProtMerge_Results'))

            result = messagebox.askyesno(
                "Launch Similarity Analysis",
                f"File: {file_name}\n"
                f"Proteins: {total_proteins}\n\n"
                f"Launch similarity analysis on this data?"
            )

            if result:
                # Only disable similarity mode button if it exists and we're in similarity mode
                similarity_btn_to_disable = None
                if (hasattr(self, 'launch_similarity_btn') and
                    self.launch_similarity_btn and
                    self.current_mode == "similarity_mode"):
                    similarity_btn_to_disable = self.launch_similarity_btn
                    similarity_btn_to_disable.config(state="disabled", text="Analyzing...")

                try:
                    # Pass self.app (MockApp in testing, actual app in production)
                    similarity_result = launch_similarity_analysis(self.root, file_path, self.app)

                    if similarity_result == "complete":
                        messagebox.showinfo("Success", "Similarity analysis completed successfully!")
                    elif similarity_result == "cancelled":
                        messagebox.showinfo("Cancelled", "Similarity analysis was cancelled.")
                    else:
                        messagebox.showwarning("No Results", "Similarity analysis completed but produced no results.")

                except Exception as e:
                    messagebox.showerror("Analysis Error", f"{ERROR_MESSAGES['SIMILARITY_ANALYSIS_FAILED']}\n{e}")

                finally:
                    # Re-enable button only if we disabled it
                    if similarity_btn_to_disable:
                        similarity_btn_to_disable.config(
                            state="normal",
                            text="🔬 Launch Similarity Analyzer"
                        )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch similarity analysis:\n{e}")

    # =============================================================================
    # RESULTS MANAGEMENT METHODS
    # =============================================================================

    def _launch_data_viewer(self):
        """Launch data viewer with file selection"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File to View",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            self._show_data_viewer(file_path)

    def _show_data_viewer(self, file_path):
        """Show enhanced Excel data viewer using dedicated module"""
        try:
            # The import is now handled at the top of the file,
            # using mock if the actual module is not found.
            # from excel_data_viewer import launch_excel_viewer # Moved to top

            # Launch the dedicated viewer with close callback
            def on_viewer_close():
                # Return focus to main window
                self.root.focus_set()
                # If we came from completion dialog, don't auto-return to main menu.
                # The completion dialog already determined the next action.
                if not self.completion_dialog:
                    self._return_to_main_menu() # Only return to main menu if viewer launched directly

            viewer = launch_excel_viewer(self.root, file_path, on_viewer_close)

            if not viewer:
                # Fallback to simple message if viewer fails
                messagebox.showerror("Error", "Could not open Excel data viewer")

        except ImportError:
            messagebox.showerror(
                "Module Missing",
                f"{ERROR_MESSAGES['MODULE_MISSING']}\n\n"
                "Please ensure excel_data_viewer.py is present in the application directory."
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open data viewer:\n{e}")

    def _reset_for_new_analysis(self):
        """Reset interface for new analysis"""
        self.file_selected = False
        self.column_selected = False
        self.input_file = None
        self.uniprot_column = None
        self.completion_dialog = None
        self.analysis_in_progress = False
    
        # Reset gene ID toggle to default state
        if hasattr(self, 'use_gene_ids') and self.use_gene_ids:
            self.use_gene_ids.set(False)
            self.current_options['use_gene_ids'] = False

        # Hide progress (in case it was visible)
        self._hide_progress()

        # Reset to new analysis mode (this also clears content)
        self._switch_to_new_analysis()

        # Explicitly clear file and column selections if widgets exist
        if hasattr(self, 'file_label'):
            self.file_label.config(text="No file selected", fg=Theme.TEXT_MUTED)

        if hasattr(self, 'column_listbox'):
            self.column_listbox.delete(0, tk.END)

        if hasattr(self, 'sheet_combo'):
            self.sheet_combo['state'] = 'disabled'
            self.sheet_combo.set('')

        self._update_status("Select file and column to continue", Theme.TEXT_MUTED)
        self._update_start_button()
    
    def _show_statistics(self):
        """Show statistics tool"""
        messagebox.showinfo("Coming Soon", "Statistics tool will be available in a future update.")

    def _search_proteins(self):
        """Show protein search tool"""
        messagebox.showinfo("Coming Soon", "Protein search tool will be available in a future update.")

    def _export_data(self):
        """Show export options"""
        messagebox.showinfo("Coming Soon", "Export tool will be available in a future update.")

    def _merge_files(self):
        """Show file merge tool"""
        messagebox.showinfo("Coming Soon", "File merge tool will be available in a future update.")

    def _compare_results(self):
        """Show results comparison tool"""
        messagebox.showinfo("Coming Soon", "Results comparison tool will be available in a future update.")

    def _show_export_options(self, output_file):
        """Show export options for completed analysis"""
        # This is a placeholder for actual export functionality
        messagebox.showinfo(
            "Export Options",
            f"Export functionality for {Path(output_file).name} will be implemented in a future update.\n\n"
            f"For now, you can find your results in the Excel file that was created."
        )
        # This will return to main menu via the caller (_show_completion)

    # =============================================================================
    # UTILITY METHODS
    # =============================================================================

    def _show_options(self):
        """Show options dialog"""
        # Sync current GUI state to options before showing modal
        self.current_options['use_gene_ids'] = self.use_gene_ids.get()
    
        modal = OptionsModal(self.root, self.current_options)
        result = modal.show()
        if result:
            self.current_options = result
            self.options_configured = True
        
            # NEW: Sync the gene ID toggle button state from modal result
            gene_ids_enabled = result.get('use_gene_ids', False)
            self.use_gene_ids.set(gene_ids_enabled)
            self._update_gene_toggle_appearance()
        
            self._update_options_summary()
            self._update_start_button()

    def _update_options_summary(self):
        """Update options summary display"""
        opts = []
        if self.current_options.get('protparam'):
            opts.append("ProtParam")
            if self.current_options.get('amino_acid'):
                opts.append("AminoAcids")
        if self.current_options.get('blast'):
            opts.append("BLAST")
        if self.current_options.get('pdb_search'):
            opts.append("PDB")

        summary = f"UniProt + {' + '.join(opts)}" if opts else "UniProt only"
    
        # Add gene ID mode indicator
        if self.current_options.get('use_gene_ids'):
            summary = f"Gene→{summary}"
    
        if self.current_options.get('safe_mode'):
            summary += " [Safe]"

        if hasattr(self, 'options_summary'):
            self.options_summary.config(text=summary)

    def _update_start_button(self):
        """Update start button state"""
        if hasattr(self, 'start_btn'):
            ready = self.file_selected and self.column_selected
            if ready:
                self.start_btn.config(
                    state="normal",
                    bg=Theme.GREEN,
                    activebackground="#45a049",
                    text="START ANALYSIS"
                )
            else:
                self.start_btn.config(
                    state="disabled",
                    bg="#4a4a4a",
                    activebackground="#4a4a4a",
                    text="START ANALYSIS"
                )

    def _update_status(self, msg, color):
        """Update status label"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=msg, fg=color)

    def _open_file(self, file_path):
        """Open file with system default application"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(str(file_path))
            elif os.name == 'posix':  # macOS and Linux
                if os.uname().sysname == 'Darwin':  # macOS
                    subprocess.run(['open', str(file_path)])
                else:  # Linux
                    subprocess.run(['xdg-open', str(file_path)])
        except Exception as e:
            messagebox.showerror("Error", f"{ERROR_MESSAGES['FILE_OPEN_ERROR']} {e}")

    def _cancel(self):
        """Cancel current operation"""
        if messagebox.askyesno("Cancel", "Cancel current operation?"):
            # If an analysis is running, we might need more sophisticated cancellation
            # For now, just exit if confirmed.
            self.root.quit()

    def _center_window(self):
        """Center window on screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - self.root.winfo_width()) // 2
        y = (self.root.winfo_screenheight() - self.root.winfo_height()) // 2
        self.root.geometry(f"+{x}+{y}")
    
    def _update_gene_toggle_appearance(self):
        """Update the gene toggle button appearance based on current state"""
        if hasattr(self, 'gene_toggle') and self.gene_toggle:
            if self.use_gene_ids.get():
                self.gene_toggle.set_nav_active(True)
                self.gene_toggle.config(text="Gene IDs ✓")
                if hasattr(self, 'column_label'):
                    self.column_label.config(text="Select column with Gene IDs:")
            else:
                self.gene_toggle.set_nav_active(False)
                self.gene_toggle.config(text="Gene IDs")
                if hasattr(self, 'column_label'):
                    self.column_label.config(text="Select column with UniProt IDs:")

            # Update status if column is selected
            if self.column_selected:
                mode_text = "Gene ID" if self.use_gene_ids.get() else "UniProt ID"
                try:
                    col_name = self.column_listbox.get(self.column_listbox.curselection()[0]).split('|')[1].strip()
                    self._update_status(f"Selected: {col_name} ({mode_text} mode)", Theme.GREEN)
                except:
                    pass


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    """Main entry point for GUI testing"""
    import logging

    # Setup basic logging for testing
    logging.basicConfig(level=logging.INFO)

    # Mock app for testing
    class MockApp:
        def run_analysis(self, input_file, sheet_name, uniprot_column_index, options, progress_callback):
            self.logger.info(f"MockApp: Running analysis with {input_file}, sheet '{sheet_name}', column {uniprot_column_index}, options {options}")
            # Simulate work
            for i in range(101):
                time.sleep(0.05) # Simulate processing time
                if i <= 40:
                    progress_callback(i, f"Fetching UniProt data... ({i}%)")
                elif i <= 65:
                    progress_callback(i, f"Performing ProtParam analysis... ({i-40}%)")
                elif i <= 85:
                    progress_callback(i, f"Running BLAST search... ({i-65}%)")
                else:
                    progress_callback(i, f"Searching PDB structures... ({i-85}%)")
            self.logger.info("MockApp: Analysis complete.")
            return "test_output.xlsx"

        def get_analysis_summary(self):
            # This is the crucial part that gets passed to CompletionDialog
            return {
                'total_proteins': 100,
                'uniprot_complete': 95,
                'uniprot_percent': 95.0,
                'protparam_complete': 90,
                'protparam_percent': 90.0,
                'blast_complete': 70,
                'blast_percent': 70.0,
                'pdb_complete': 50,
                'pdb_percent': 50.0
            }
        def __init__(self):
            self.logger = logging.getLogger(__name__)

    app = MockApp()
    gui = ProtMergeGUI(app)
    gui.run()


if __name__ == "__main__":
    main()