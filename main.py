import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib.parse
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
import threading
from datetime import datetime
import random
from typing import Dict, List, Optional
from dataclasses import dataclass
from pathlib import Path
import json

# Constants
CONFIG_FILE = "config.json"
LOG_FILE = "whatsapp_automation.log"
TEMPLATE_FILE = "message_template.txt"
PROFILE_DIR = "whatsapp_profile"


@dataclass
class Contact:
    """Data class for contact information"""
    phone: str
    name: Optional[str] = None


@dataclass
class AutomationConfig:
    """Data class for automation configuration"""
    min_delay: float = 30.0
    max_delay: float = 120.0
    headless: bool = True
    save_reports: bool = True
    random_order: bool = False
    dark_mode: bool = False


class AutomationResult:
    """Class to handle automation results"""

    def __init__(self):
        self.successful: List[str] = []
        self.failed: List[Dict[str, str]] = []
        self.messages_sent: Dict[str, str] = {}

    def add_success(self, phone: str, message: str):
        self.successful.append(phone)
        self.messages_sent[phone] = message

    def add_failure(self, contact: Contact):
        self.failed.append({
            'name': contact.name,
            'phone': contact.phone
        })


class WhatsAppAutomationGUI:
    def __init__(self):
        """Initialize GUI application"""
        self.root = tk.Tk()
        self.setup_window()
        self.load_config()
        self.init_variables()
        self.setup_styles()
        self.create_widgets()
        self.setup_bindings()
        self.excel_path = None
        self.automation = None
        self.df = None
        self.phone_column = None
        self.name_column = None
        self.columns = []
        self.status_column_index = None
        self.is_paused = False
        self.is_running = False
        self.start_time = None
        self.total: int = 0
        self.processed_count = 0
        self.message_variations = []

    def setup_window(self):
        """Configure the main window properties"""
        self.root.title("WhatsApp Automation")
        self.root.geometry("1000x800")
        self.root.configure(bg='#f0f2f5')

        # Make the window resizable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def load_config(self):
        """Load saved configuration from file"""
        self.config = AutomationConfig()
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config_data = json.load(f)
                    self.config = AutomationConfig(**config_data)
            except Exception as e:
                logging.error(f"Error loading config: {e}")

    def save_config(self):
        """Save current configuration to file"""
        try:
            config_data = {
                'min_delay': self.min_delay.get(),
                'max_delay': self.max_delay.get(),
                'headless': self.headless_var.get(),
                'save_reports': self.save_reports_var.get(),
                'random_order': self.random_order_var.get(),
                'dark_mode': self.dark_mode_var.get(),
            }
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config_data, f, indent=4)

        except Exception as e:
            logging.error(f"Error saving config: {e}")

    def init_variables(self):
        """Initialize application variables"""
        self.excel_path: Optional[str] = None
        self.automation: Optional[WhatsAppAutomation] = None
        self.df: Optional[pd.DataFrame] = None
        self.phone_column: Optional[str] = None
        self.name_column: Optional[str] = None
        self.columns: List[str] = []
        self.status_column_index: Optional[int] = None
        self.is_paused: bool = False
        self.is_running: bool = False
        self.start_time: Optional[float] = None
        self.total: int = 0
        self.processed_count: int = 0
        self.message_variations: List[str] = []
    def setup_styles(self):
        """Configure custom styles for widgets"""
        if self.config.dark_mode:
            style = ttk.Style()
            style.theme_use('clam')
            self.root.config(bg='#121212')

            style.configure("TFrame",
                            background='#1e1e1e',
                            borderwidth=0)

            style.configure("TLabel",
                            background='#1e1e1e',
                            foreground='#e5e5e5',
                            font=('Helvetica', 10))

            style.configure("TLabelframe",
                            background='#1e1e1e')

            style.configure("TLabelframe.Label",
                            background='#1e1e1e',
                            foreground='#e5e5e5')

            style.configure("TButton",
                            padding=10,
                            background='#1e1e1e',
                            foreground='#e5e5e5',
                            font=('Helvetica', 10))
            style.map("TButton",
                      background=[("active", "#2e2e2e"), ("disabled", "#222222")],
                      foreground=[("active", "#9a9a99"), ("disabled", "#3a3b3b")])

            style.configure("Custom.TButton",
                            background='#1e1e1e',
                            foreground='#e5e5e5',
                            padding=1,
                            width=17,
                            height=3,
                            font=('Helvetica', 10))
            style.map("Custom.TButton",
                      background=[("active", "#2e2e2e"), ("disabled", "#222222")],
                      foreground=[("active", "#9a9a99"), ("disabled", "#3a3b3b")])

            style.configure("TEntry",
                            foreground="#e5e5e5",
                            fieldbackground="#2b2b2b",
                            background="#2b2b2b")

            style.configure("TCheckbutton",
                            foreground="#e5e5e5",
                            background="#1e1e1e")
            style.map("TCheckbutton",
                      background=[("active", "#2e2e2e")],
                      foreground=[("active", "#9a9a99")])

            style.configure("Horizontal.TProgressbar",
                            thickness=30,
                            troughcolor="#2b2b2b",
                            background="#4caf50",
                            barcolor="#4caf50")

            style.configure("Treeview",
                            rowheight=25,
                            background="#2b2b2b",
                            foreground="#e5e5e5",
                            fieldbackground="#2b2b2b",
                            font=('Helvetica', 10))

            style.configure("Treeview.Heading",
                            background="2b2b2b",
                            foreground="#e5e5e5")

            style.configure("Vertical.TScrollbar",
                            background="#1e1e1e",
                            troughcolor="#2b2b2b",
                            arrowcolor="#e5e5e5")
            style.configure("Horizontal.TScrollbar",
                            background="#1e1e1e",
                            troughcolor="#2b2b2b",
                            arrowcolor="#e5e5e5")

        else:
            style = ttk.Style()
            style.configure("TButton",
                            padding=10,
                            font=('Helvetica', 10))
            style.configure("Custom.TButton",
                            padding=1,
                            width=17,
                            height=3,
                            font=('Helvetica', 10))
            style.configure("TFrame",
                            background='#f0f0f0')
            style.configure("Treeview",
                            rowheight=25,
                            font=('Helvetica', 10))
            style.configure("TLabel",
                            background='#f0f2f5',
                            font=('Helvetica', 10))

    def create_widgets(self):
        """Create and arrange all GUI widgets with a scalable grid"""
        self.main_container = ttk.Frame(self.root, style="TFrame")
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Configure Main Container Rows
        # Row 0: File (Fixed height)
        # Row 1: Middle (Message/Settings - Scalable)
        # Row 2: Preview (Treeview - Most Scalable)
        # Row 3: Footer (Controls/Progress - Fixed height)
        self.main_container.columnconfigure(0, weight=1)
        self.main_container.rowconfigure(1, weight=2)
        self.main_container.rowconfigure(2, weight=4)

        # 1. TOP
        self.create_file_section(self.main_container).grid(row=0, column=0, sticky="ew", pady=(0, 10))

        # 2. MIDDLE (Side-by-Side)
        self.middle_container = ttk.Frame(self.main_container, style="TFrame")
        self.middle_container.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        self.middle_container.columnconfigure(0, weight=3)  # Message gets more width
        self.middle_container.columnconfigure(1, weight=1)  # Settings gets less
        self.middle_container.rowconfigure(0, weight=1)

        self.create_message_section(self.middle_container)
        self.create_settings_section(self.middle_container)

        # 3. PREVIEW
        self.create_preview_section(self.main_container).grid(row=2, column=0, sticky="nsew")

        # 4. FOOTER
        self.footer_container = ttk.Frame(self.main_container, style="TFrame")
        self.footer_container.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.create_control_section(self.footer_container)
        self.create_progress_section(self.footer_container)

    def create_file_section(self, parent):
        """Create file selection section"""
        file_frame = ttk.LabelFrame(parent, style="TLabelframe", text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 5))

        self.file_label = ttk.Label(file_frame, style="TLabel", text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=(0, 10))

        select_button = ttk.Button(
            file_frame,
            text="Select Excel File",
            command=self.select_file,
            style="TButton"
        )
        select_button.pack(side=tk.RIGHT)
        return file_frame

    def create_message_section(self, parent):
        """Create message input section"""
        message_frame = ttk.LabelFrame(parent, style="TLabelframe", text="Message Templates", padding=10)
        message_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # Container for Text + Scrollbar
        text_container = ttk.Frame(message_frame, style="TFrame")
        text_container.pack(fill=tk.BOTH, expand=True)

        self.message_text = tk.Text(text_container, height=5, font=('Helvetica', 10),
                                    bg='#2b2b2b' if self.config.dark_mode else 'white',
                                    fg='white' if self.config.dark_mode else 'black',
                                    undo=True, wrap=tk.WORD)

        text_vsb = ttk.Scrollbar(text_container, orient="vertical", command=self.message_text.yview)
        self.message_text.configure(yscrollcommand=text_vsb.set)

        text_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.message_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bottom container for labels/buttons to keep them from stretching awkwardly
        btn_container = ttk.Frame(message_frame, style="TFrame")
        btn_container.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(btn_container, text="Variables: {name}", font=('Helvetica', 9, 'italic')).pack(side=tk.LEFT)
        self.save_template_button = ttk.Button(btn_container, text="Save Template", style="Custom.TButton",
                                               command=self.save_template)
        self.save_template_button.pack(side=tk.RIGHT, padx=2)
        self.load_template_button = ttk.Button(btn_container, text="Load Template", style="Custom.TButton",
                                               command=self.load_template)
        self.load_template_button.pack(side=tk.RIGHT, padx=2)

    def create_settings_section(self, parent):
        """Create settings section"""
        settings_frame = ttk.LabelFrame(parent, style="TLabelframe", text="Settings", padding=10)
        settings_frame.grid(row=0, column=1, sticky="nsew")

        # Delay settings
        delay_frame = ttk.Frame(settings_frame,  style="TFrame")
        delay_frame.pack(fill=tk.X)

        ttk.Label(delay_frame, style="TLabel", text="Delay Range (seconds):").pack(side=tk.LEFT)
        self.min_delay = ttk.Entry(delay_frame, style="TEntry", width=5)
        self.min_delay.insert(0, str(self.config.min_delay))
        self.min_delay.pack(side=tk.LEFT, padx=5)

        ttk.Label(delay_frame, style="TLabel", text="to").pack(side=tk.LEFT)
        self.max_delay = ttk.Entry(delay_frame, style="TEntry", width=5)
        self.max_delay.insert(0, str(self.config.max_delay))
        self.max_delay.pack(side=tk.LEFT, padx=5)

        # Additional settings
        self.headless_var = tk.BooleanVar(value=self.config.headless)
        self.headless_check = ttk.Checkbutton(
            settings_frame,
            style='TCheckbutton',
            text="Run in headless mode (after initial login)",
            variable=self.headless_var,
            command=self.save_config
        )
        self.headless_check.pack(anchor=tk.W)

        self.save_reports_var = tk.BooleanVar(value=self.config.save_reports)
        self.save_reports_check = ttk.Checkbutton(
            settings_frame,
            style='TCheckbutton',
            text="Save detailed reports",
            variable=self.save_reports_var,
            command=self.save_config
        )
        self.save_reports_check.pack(anchor=tk.W)

        self.random_order_var = tk.BooleanVar(value=self.config.random_order)
        self.random_order_check = ttk.Checkbutton(
            settings_frame,
            text="Pick contacts in random order",
            style='TCheckbutton',
            variable=self.random_order_var,
            command=self.save_config
        )
        self.random_order_check.pack(anchor=tk.W)

        self.dark_mode_var = tk.BooleanVar(value=self.config.dark_mode)
        self.dark_mode_check = ttk.Checkbutton(
            settings_frame,
            text="Dark mode (requires relaunch)",
            style='TCheckbutton',
            variable=self.dark_mode_var,
            command=self.save_config
        )
        self.dark_mode_check.pack(anchor=tk.W)

    def create_preview_section(self, parent):
        """Create data preview section with scrollbars"""
        preview_frame = ttk.LabelFrame(parent, style='TLabelframe', text="Data Preview", padding=10)

        # Configure grid weights so the Treeview (0,0) expands, but scrollbars don't
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Create Treeview
        self.tree = ttk.Treeview(
            preview_frame,
            style="Treeview",
            show='headings',
            selectmode='browse',
            height=7
        )

        # Create Scrollbars
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid them:
        # Treeview in the main cell
        self.tree.grid(column=0, row=0, sticky='nsew')

        # Vertical bar on the right
        vsb.grid(column=1, row=0, sticky='ns')

        # Horizontal bar on the bottom
        hsb.grid(column=0, row=1, sticky='ew')

        # Status colors
        self.tree.tag_configure('success', background='#2e4a2e' if self.config.dark_mode else '#90EE90')
        self.tree.tag_configure('error', background='#4a2e2e' if self.config.dark_mode else '#FFB6C1')
        self.tree.tag_configure('pending', background='#4a432e' if self.config.dark_mode else '#FFE4B5')

        return preview_frame

    def create_control_section(self, parent):
        """Create control buttons section"""
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_button = ttk.Button(
            control_frame,
            text="Start Automation",
            command=self.start_automation,
            style="TButton"
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.pause_button = ttk.Button(
            control_frame,
            text="Pause",
            command=self.toggle_pause,
            state=tk.DISABLED,
            style="TButton"
        )
        self.pause_button.pack(side=tk.LEFT, padx=5)

        self.retry_button = ttk.Button(
            control_frame,
            text="Retry Failed",
            command=self.retry_failed,
            state=tk.DISABLED,
            style="TButton"
        )
        self.retry_button.pack(side=tk.LEFT, padx=5)

    def create_progress_section(self, parent):
        """Create progress tracking section"""
        progress_frame = ttk.LabelFrame(parent, style="TLabelframe", text="Progress", padding=10)
        progress_frame.pack(fill=tk.X, pady=(10, 0))

        self.progress_label = ttk.Label(progress_frame, style="TLabel", text="")
        self.progress_label.pack()

        self.time_label = ttk.Label(progress_frame, style="TLabel", text="")
        self.time_label.pack()

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            style="Horizontal.TProgressbar",
            length=300,
            mode='determinate'
        )
        self.progress_bar.pack()

        self.status_label = ttk.Label(
            progress_frame,
            style="TLabel",
            text="",
            wraplength=700
        )
        self.status_label.pack(pady=10)


    def setup_bindings(self):
        """Setup event bindings"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind('<Control-q>', lambda e: self.on_closing())
        self.root.bind('<F5>', lambda e: self.update_preview())

    def on_closing(self):
        """Handle application closing"""
        if self.is_running:
            if not messagebox.askyesno("Confirm Exit",
                                       "Automation is running. Are you sure you want to exit?"):
                return
        self.save_config()
        self.root.destroy()

    def save_template(self):
        """Save message template with non-blocking visual feedback"""
        self.save_template_button.config(state=tk.DISABLED)
        template = self.message_text.get("1.0", tk.END).strip()
        try:
            with open(TEMPLATE_FILE, 'w', encoding="utf-8") as msg:
                msg.write(template)
            self.save_template_button.config(text="Saved!")
        except Exception as e:
            logging.error(f'Error saving template: {e}')
            self.save_template_button.config(text="Error Saving")
        finally:
            # Reset button after 1.5 seconds without freezing
            self.root.after(1500, lambda: self.save_template_button.config(
                text="Save Template",
                state=tk.NORMAL
            ))

    def load_template(self):
        """Load saved message template and replace current text"""
        self.load_template_button.config(state=tk.DISABLED)
        try:
            with open(TEMPLATE_FILE, 'r', encoding="utf-8") as msg:
                template = msg.read()
                self.message_text.delete("1.0", tk.END)
                self.message_text.insert("1.0", template)
            self.load_template_button.config(text="Template Loaded!")

        except FileNotFoundError:
            print("No template file found.")
            self.load_template_button.config(text="No File Found")
        except Exception as e:
            print(f'Error loading template: {e}')
            self.load_template_button.config(text="Error Loading")
        finally:
            # Reset button state after 1 second without freezing the UI
            self.root.after(1000, lambda: self.load_template_button.config(
                text="Load Template",
                state=tk.NORMAL
            ))

    def detect_phone_column(self, df):
        """Detect the column containing phone numbers"""
        phone_patterns = [
            r'^\+?\d{10,}$',  # Basic phone number pattern
            r'^(?:\+\d{1,3}|0)\d{9,}$',  # International format
            r'phone|mobile|contact|tel',  # Common column names
        ]

        for col in df.columns:
            # Check column name
            if any(re.search(pattern, str(col).lower()) for pattern in phone_patterns[2:]):
                return col

            # Check column values
            if df[col].astype(str).str.match('|'.join(phone_patterns[:2])).any():
                return col

        return None

    def detect_name_column(self, df):
        """Detect the column containing names"""
        name_patterns = [
            r'name|full\s*name|customer|client',  # Common column names
            r'^[A-Za-z\s]{2,}$'  # Basic name pattern
        ]

        for col in df.columns:
            # Check column name
            if re.search(name_patterns[0], str(col).lower()):
                return col

            # Check column values
            if df[col].astype(str).str.match(name_patterns[1]).any():
                return col

        return None

    def update_preview(self):
        """Update the data preview in the Treeview"""
        # Clear existing items
        self.tree.delete(*self.tree.get_children())
        for col in self.tree['columns']:
            self.tree.heading(col, text='')

        if self.df is not None:
            # Add status columns
            self.df['Status'] = ''
            self.df['Details'] = ''

            # Configure columns
            columns = list(self.df.columns)
            self.columns = columns  # Store columns for later use
            self.tree['columns'] = columns
            self.status_column_index = columns.index('Status')

            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)

            # Add data
            for i, row in self.df.head(100).iterrows():
                self.tree.insert('', 'end', values=list(row))

            # Highlight detected columns
            if self.phone_column:
                self.tree.heading(self.phone_column, text=f"📱 {self.phone_column}")
            if self.name_column:
                self.tree.heading(self.name_column, text=f"👤 {self.name_column}")

    def select_file(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                self.excel_path = file_path
                self.file_label.config(text=f"Selected: {os.path.basename(file_path)}")

                # Load Excel file
                self.df = pd.read_excel(file_path)

                # Detect columns
                self.phone_column = self.detect_phone_column(self.df)
                self.name_column = self.detect_name_column(self.df)

                if not self.phone_column:
                    messagebox.showwarning("Warning", "Could not automatically detect phone number column!")

                # Update preview
                self.update_preview()

                # Show detected columns with counts
                total_rows = len(self.df)
                detection_text = "Detected columns:\n"

                if self.phone_column:
                    phone_count = self.df[self.phone_column].notna().sum()
                    detection_text += f"Phone Numbers: {phone_count}/{total_rows} detected\n"
                else:
                    detection_text += "Phone Numbers: Not detected\n"

                if self.name_column:
                    name_count = self.df[self.name_column].notna().sum()
                    detection_text += f"Names: {name_count}/{total_rows} detected"
                else:
                    detection_text += "Names: Not detected"

                self.status_label.config(text=detection_text)

            except Exception as e:
                messagebox.showerror("Error", f"Error loading Excel file: {str(e)}")

    def toggle_pause(self):
        """Toggle pause/resume state"""
        self.is_paused = not self.is_paused
        self.pause_button.config(text="Resume" if self.is_paused else "Pause")

        # Store the original status text so we can restore it later
        pre_pause_text = self.status_label.cget("text")
        pause_message = "\nAutomation PAUSED. Click Resume to continue."
        resume_message = "\n Resuming automation..."

        if self.is_paused:
            self.status_label.config(text=pre_pause_text + pause_message)
        else:
            self.status_label.config(text=resume_message)

    def retry_failed(self):
        """Retry failed messages"""
        failed_contacts = []
        tree_col = ('name', 'phone', 'Status', 'Details')
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if values[self.status_column_index] == 'Failed ✗':
                row_dict = dict(zip(tree_col, values))
                failed_contacts.append(row_dict)

        if failed_contacts:
            if messagebox.askyesno("Retry Failed", f"Retry sending {len(failed_contacts)} failed messages?"):
                self.start_automation(retry_mode=True, retry_contacts=failed_contacts)
        else:
            messagebox.showinfo("Retry Failed", "No failed messages to retry.")

    def update_estimated_time(self, processed: int, total: int):
        """Update estimated time remaining."""

        if processed > round(total/3):
            # Avoid division by zero and provide a better first estimate
            if self.start_time and processed > 0:
                elapsed_time = time.time() - self.start_time
                avg_time_per_item = elapsed_time / processed

        else:
            # Use user-configured delay range as initial estimate
            try:
                min_d = float(self.min_delay.get())
                max_d = float(self.max_delay.get())
                avg_time_per_item = (min_d + max_d) / 2
            except Exception:
                avg_time_per_item = 60 * 5  # default 5 min

        remaining_items = max(total - processed, 0)
        estimated_remaining = remaining_items * avg_time_per_item

        # Convert to human-friendly format
        def format_duration(seconds: float) -> str:
            if seconds < 60:
                return f"{int(seconds)} seconds"
            elif seconds < 3600:
                return f"{seconds / 60:.1f} minutes"
            else:
                return f"{seconds / 3600:.1f} hours"

        time_str = format_duration(estimated_remaining)

        # Calculate ETA
        eta_timestamp = time.time() + estimated_remaining
        eta_str = datetime.fromtimestamp(eta_timestamp).strftime('%I:%M:%S %p').lstrip('0')

        self.time_label.config(
            text=f"Estimated time remaining: {time_str} ({eta_str})"
        )

    def update_status_in_tree(self, phone, status, details):
        """Update status and highlight the current active row"""
        for item in self.tree.get_children():
            values = list(self.tree.item(item)['values'])
            if str(values[self.columns.index(self.phone_column)]) == str(phone):
                values[self.status_column_index] = status
                values[self.status_column_index + 1] = details

                # Apply tags for coloring
                tag = 'success' if 'Success' in status else 'error'
                self.tree.item(item, values=values, tags=(tag,))

                self.tree.see(item)  # Auto-scroll to this item
                self.tree.selection_set(item)  # Visually select the active row
                break
        self.root.update()

    def update_status_tracking(self, results, completed, output_path=None):
        """Update status and details for a specific contact to exel"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status_messages = []
            detailed_results = []

            # Prepare detailed results data
            for phone in results['successful']:
                detailed_results.append({
                    'Phone': phone,
                    'Status': 'Success',
                    'Message_Sent': results['messages_sent'].get(phone, ''),
                    'Timestamp': timestamp,
                    'Details': 'Message sent successfully'
                })

            for failed in results['failed']:
                detailed_results.append({
                    'Phone': failed['phone'],
                    'Status': 'Failed',
                    'Message_Sent': '',
                    'Timestamp': timestamp,
                    'Details': f"Failed to send message: {str(failed.get('details', 'Unknown error'))}"
                })

            # Convert to DataFrame for easier manipulation
            results_df = pd.DataFrame(detailed_results)

            # Update Excel status file
            if self.df is not None:
                if not output_path and hasattr(self, 'excel_path'):
                    base_name, ext = os.path.splitext(self.excel_path)
                    output_path = f"{base_name}_report{ext}"

                if output_path:
                    # Update existing DataFrame with status information
                    for _, row in results_df.iterrows():
                        mask = self.df[self.phone_column].astype(str) == str(row['Phone'])
                        if any(mask):
                            self.df.loc[mask, 'Status'] = row['Status']
                            self.df.loc[mask, 'Details'] = row['Details']
                            self.df.loc[mask, 'Message_Sent'] = row['Message_Sent']
                            self.df.loc[mask, 'Timestamp'] = row['Timestamp']

                    # Save updated DataFrame
                    self.df.to_excel(output_path, sheet_name='Detailed Results', index=False)
                    status_messages.append(f"Excel status updated: {output_path}")

        except Exception as e:
            error_msg = f"Error updating status tracking: {str(e)}"
            logging.error(error_msg)
            return error_msg

        try:
            # Generate comprehensive report if completed
            if completed:
                # Prepare summary data
                successful = self.df[self.df['Status'] == 'Success']
                failed = self.df[self.df['Status'] == 'Failed']
                total = len(successful) + len(failed)
                success_rate = (len(successful) / total * 100) if total > 0 else 0

                summary_data = pd.DataFrame({
                    'Metric': ['Total Messages', 'Successful', 'Failed', 'Success Rate'],
                    'Value': [total, len(successful), len(failed), f"{success_rate:.1f}%"]
                })

                # Prepare unsuccessful contacts data
                unsuccessful_contacts = failed[[self.phone_column, 'Status', 'Details' , 'Timestamp']].copy()
                if 'Name' in self.df.columns:
                    unsuccessful_contacts['Name'] = self.df.loc[failed.index, 'Name'].values
                    cols = ['Name', self.phone_column, 'Status', 'Details', 'Timestamp']
                    unsuccessful_contacts = unsuccessful_contacts[cols]

                # Save all sheets to Excel report
                with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    summary_data.to_excel(writer, sheet_name='Summary', index=False)
                    unsuccessful_contacts.to_excel(writer, sheet_name='Unsuccessful Contacts', index=False)

                status_messages.append(f"Report generated: {output_path}")

                report_message = (
                    f"Automation completed!\n"
                    f"Report generated: {output_path}\n"
                    f"Total messages: {total}\n"
                    f"Successful: {len(successful)}\n"
                    f"Failed: {len(failed)}\n"
                    f"Success rate: {success_rate:.1f}%\n"
                    f"Report saved as: {output_path}"
                )

                status_messages.append(report_message)
                self.status_label.config(text=f"{report_message}")

        except Exception as e:
            error_msg = f"Error generating comprehensive report: {str(e)}"
            logging.error(error_msg)
            return error_msg

        logging.info("; ".join(status_messages))
        return "; ".join(status_messages)

    def update_excel_status(self, phone, status, details, message_sent=None):
        """Update status in Excel file for a single contact"""
        results = {
            'successful': [phone] if status == 'Success' else [],
            'failed': [{'phone': phone, 'name': None, 'details': details}] if status == 'Failed' else [],
            'messages_sent': {phone: message_sent} if message_sent else {}
        }
        completed = self.total - self.processed_count == 1
        return self.update_status_tracking(results, completed)

    def start_automation(self, retry_mode=False, retry_contacts=None):
        """Start the WhatsApp automation process"""
        if not self.excel_path or not self.phone_column:
            messagebox.showerror("Error", "Please select an Excel file and ensure phone numbers are detected!")
            return

        # Get message variations
        messages = self.message_text.get("1.0", tk.END).strip().split('\n\n\n\n')
        messages = [msg for msg in messages if msg.strip()]
        if not messages:
            messagebox.showerror("Error", "Please enter at least one message variation!")
            return

        try:
            min_delay = float(self.min_delay.get())
            max_delay = float(self.max_delay.get())
            if min_delay > max_delay:
                messagebox.showerror("Error", "Minimum delay should be less than maximum delay!")
                return
        except ValueError:
            messagebox.showerror("Error", "Please enter valid delay values!")
            return

        self.message_variations = messages
        self.is_running = True
        self.is_paused = False
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL)
        self.retry_button.config(state=tk.DISABLED)
        self.headless_check.config(state=tk.DISABLED)
        self.save_reports_check.config(state=tk.DISABLED)
        self.dark_mode_check.config(state=tk.DISABLED)
        self.min_delay.config(state=tk.DISABLED)
        self.max_delay.config(state=tk.DISABLED)

        # Reset progress
        self.progress_label.config(text=f"Progress: (0%)")
        self.time_label.config(text="")
        self.progress_bar['value'] = 0
        self.status_label.config(text="Starting automation...")

        # Initialize automation
        self.automation = WhatsAppAutomation(
            self.excel_path,
            phone_column=self.phone_column,
            name_column=self.name_column,
            gui=self,
            config=self.config,
            min_delay=min_delay,
            max_delay=max_delay
        )

        # Run automation in a separate thread
        thread = threading.Thread(
            target=self.run_automation,
            args=(messages, retry_mode, retry_contacts)
        )
        thread.start()

    def run_automation(self, message, retry_mode=False, retry_contacts=None):
        """Run the automation process"""
        try:
            self.status_label.config(text="Starting automation process...")
            self.start_time = time.time()
            self.processed_count = 0

            # Update GUI state
            self.is_running = True
            self.start_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.NORMAL)
            self.headless_check.config(state=tk.DISABLED)
            self.save_reports_check.config(state=tk.DISABLED)
            self.dark_mode_check.config(state=tk.DISABLED)

            if retry_mode and retry_contacts:
                self.total = len(retry_contacts)
                self.progress_bar['maximum'] = self.total
            else:
                self.total = len(self.df)
                self.progress_bar['maximum'] = self.total

            if not hasattr(self, 'automation') or self.automation is None:
                self.automation = WhatsAppAutomation(
                    self.excel_path,
                    phone_column=self.phone_column,
                    name_column=self.name_column,
                    gui=self
                )

            def update_progress():
                """Update progress bar and labels"""
                if not self.is_running:
                    return

                self.processed_count += 1
                self.progress_bar['value'] = self.processed_count
                self.progress_label.config(
                    text=f"Progress: {self.processed_count}/{self.total} ({(self.processed_count / self.total * 100):.1f}%)"
                )
                self.update_estimated_time(self.processed_count, self.total)
                self.root.update()

            # Process the contacts
            if retry_mode and retry_contacts:
                results = self.automation.process_numbers(
                    message,
                    retry_mode=True,
                    retry_contacts=retry_contacts,
                    progress_callback=update_progress
                )
            else:
                results = self.automation.process_numbers(
                    message,
                    progress_callback=update_progress
                )

        except Exception as e:
            print(f"Error in automation: {str(e)}")  # Debug print
            logging.error(f"Error in automation: {str(e)}")
            self.status_label.config(text=f"Error: {str(e)}")
        finally:
            self.is_running = False
            self.start_button.config(state=tk.NORMAL)
            self.pause_button.config(state=tk.DISABLED)
            self.retry_button.config(state=tk.NORMAL)
            self.headless_check.config(state=tk.NORMAL)
            self.save_reports_check.config(state=tk.NORMAL)
            self.dark_mode_check.config(state=tk.NORMAL)


class WhatsAppAutomation:
    """Core automation class with enhanced error handling and reporting"""

    def __init__(self, excel_path: str, phone_column: str,
                 name_column: Optional[str] = None,
                 gui: Optional[WhatsAppAutomationGUI] = None,
                 config: Optional[AutomationConfig] = None,
                 min_delay=10, max_delay=30):
        self.excel_path = Path(excel_path)
        self.phone_column = phone_column
        self.name_column = name_column
        self.gui = gui
        self.config = config or AutomationConfig()
        self.driver = None
        self.user_data_dir = Path.cwd() / PROFILE_DIR
        self.min_delay = min_delay
        self.max_delay = max_delay
        self.setup_logging()

    def setup_logging(self):
        """Configure logging"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filename='whatsapp_automation.log'
        )

    def create_driver_options(self, headless=False):
        """Create Chrome options with the specified settings for stable automation."""
        options = webdriver.ChromeOptions()

        # Use persistent session data (so WhatsApp Web stays logged in)
        options.add_argument(f"user-data-dir={self.user_data_dir}")

        # Common stability/performance flags
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")

        # Prevent Chrome “automation” banners and crashes on long sessions
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        # Add a common user-agent to avoid some UI loading issues
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )

        # Headless mode tweaks
        if headless:
            options.add_argument("--headless=new")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--disable-gpu")

        return options

    def is_logged_in(self, driver):
        """Check if WhatsApp Web is logged in"""
        try:
            WebDriverWait(driver, 300).until(
                lambda d: d.find_elements(By.XPATH, '//*[@id="side"]') or
                          d.find_elements(By.CSS_SELECTOR, 'canvas')
            )
            return bool(driver.find_elements(By.XPATH, '//*[@id="side"]'))
        except:
            return False

    def load_data(self):
        """Load data from Excel file"""
        try:
            df = pd.read_excel(self.excel_path)
            data = []
            for _, row in df.iterrows():
                item = {
                    'phone': str(row[self.phone_column]),
                    'name': str(row[self.name_column]) if self.name_column else None
                }
                data.append(item)
            return data
        except Exception as e:
            logging.error(f"Error loading Excel file: {str(e)}")
            raise

    def check_login_status(self):
        """Check if WhatsApp Web is already logged in"""
        try:
            # Wait for either QR code or main chat list to appear
            WebDriverWait(self.driver, 300).until(
                lambda driver: driver.find_elements(By.CSS_SELECTOR, 'canvas') or
                               driver.find_elements(By.XPATH, '//*[@id="side"]')
            )

            # If chat list is found, we're already logged in
            if self.driver.find_elements(By.XPATH, '//*[@id="side"]'):
                logging.info("Already logged into WhatsApp Web")
                return True
            else:
                logging.info("Please scan QR code to log in")
                return False

        except Exception as e:
            logging.error(f"Error checking login status: {str(e)}")
            return False

    def wait_for_login(self):
        """Wait for user to complete login process"""
        try:
            WebDriverWait(self.driver, 300).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="side"]'))
            )
            logging.info("Successfully logged into WhatsApp Web")
            return True
        except Exception as e:
            logging.error(f"Login timeout or error: {str(e)}")
            return False

    def close_popups(self):
        """Close any WhatsApp Web popups or banners."""
        try:
            while True:
                popup_closed = False

                # Close any dialog popups
                try:
                    popup_buttons = self.driver.find_elements(
                        By.XPATH, "//div[@role='dialog']//button"
                    )
                    for btn in popup_buttons:
                        label = (btn.text or "").strip().lower()
                        if label:  # Only click visible, labeled buttons
                            btn.click()
                            logging.info(f"Closed popup dialog with button: '{label}'")
                            time.sleep(0.3)
                            popup_closed = True
                except:
                    pass

                # If no popups or banners were closed, break out of loop
                if not popup_closed:
                    break

        except Exception as e:
            logging.warning(f"Error while closing popups: {e}")

    def setup_driver(self):
        """Initialize and configure Chrome WebDriver"""
        try:
            # First check login status with a temporary visible browser
            logging.info("Checking login status...")
            if self.gui:
                self.gui.status_label.config(text="Checking WhatsApp login status...")

            temp_options = self.create_driver_options(headless=False)
            temp_driver = webdriver.Chrome(options=temp_options)
            temp_driver.get("https://web.whatsapp.com")

            logged_in = self.is_logged_in(temp_driver)
            temp_driver.quit()

            # Create the actual driver based on login status
            if logged_in:
                logging.info("Already logged in - starting in headless mode")
                if self.gui:
                    self.gui.status_label.config(text="Already logged in - running in headless mode")
                options = self.create_driver_options(headless=self.config.headless)
            else:
                logging.info("Not logged in - starting in visible mode")
                if self.gui:
                    self.gui.status_label.config(text="Please scan the QR code to log in")
                options = self.create_driver_options(headless=False)

            self.driver = webdriver.Chrome(options=options)
            self.driver.get("https://web.whatsapp.com")

            if not logged_in:
                if not self.wait_for_login():
                    raise Exception("Login timeout or failure")
                if self.gui:
                    self.gui.status_label.config(text="Successfully logged in to WhatsApp Web")

            # Additional verification for headless mode
            if logged_in:
                success = WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="side"]'))
                )
                if success:
                    logging.info("Successfully initialized headless mode")
                    return True

        except Exception as e:
            logging.error(f"Error setting up WebDriver: {str(e)}")
            if self.driver:
                self.driver.quit()
            raise

    def cooldown(self, cooldown_time, name):
        """Display cooldown time in h:m:s format, pause-aware."""
        self.remaining = cooldown_time

        while self.remaining > 0:
            # Pause handling
            if self.gui and self.gui.is_paused:
                time.sleep(0.5)
                continue

            # Format remaining time
            hours = round(self.remaining) // 3600
            minutes = (round(self.remaining) % 3600) // 60
            seconds = round(self.remaining) % 60

            parts = []
            if hours > 0:
                parts.append(f"{hours:01}h")
            if minutes > 0:
                parts.append(f"{minutes:01}m")
            if seconds > 0:
                parts.append(f"{seconds:01}s")

            time_str = " ".join(parts)
            if self.gui:
                self.gui.status_label.config(text=f"Sending next message to {name} in: {time_str}")

            time.sleep(1)  # Tick down every second
            self.remaining -= 1  # Reduce only when not paused

        if self.gui:
            self.gui.status_label.config(text="")

    def send_message(self, contact, message_template):
        """Send message to a specific contact"""
        try:
            # Replace {name} placeholder with contact name if available
            message = message_template.replace("{name}", contact['name']) if contact['name'] else message_template

            # Format the URL
            url = f"https://web.whatsapp.com/send?phone={contact['phone']}&text={urllib.parse.quote(message)}"
            self.driver.get(url)

            # Define the XPaths
            error_xpaths = [
                "//*[contains(text(), 'Phone number shared via url is invalid.')]",
                "//div[@aria-label='Phone number shared via url is invalid.']"
            ]
            chat_box_xpath = "//div[@contenteditable='true']"

            # Wait for chat or error
            try:
                WebDriverWait(self.driver, 20).until(
                    lambda d: any(d.find_elements(By.XPATH, xp) for xp in error_xpaths)
                              or d.find_elements(By.XPATH, chat_box_xpath)
                )
            except Exception:
                logging.error(f"Timeout waiting for chat or error for {contact['phone']}")
                print(f"Timeout waiting for chat or error for {contact['phone']}")
                return False, None, "Timeout waiting for chat or error"

            # Handle invalid number
            for xp in error_xpaths:
                if self.driver.find_elements(By.XPATH, xp):
                    print(xp)
                    logging.warning(f"Invalid phone number for {contact['phone']} (matched: {xp}). Skipping.")
                    print(f"Invalid phone number for {contact['phone']}. Skipping.")
                    return False, None, "Invalid phone number"

            # Wait for chat input to be ready
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, chat_box_xpath))
                )
            except Exception:
                logging.error(f"Chat input not loaded for {contact['phone']}")
                print(f"Chat input not loaded for {contact['phone']}")
                return False, None, "Chat input not loaded"

            # Find send button (robust multi-check)
            possible_xpaths = [
                "//button[@aria-label='Send']",
                "//button[@data-icon='wds-ic-send-filled']",
                "//span[@data-icon='wds-ic-send-filled']/ancestor::button",
                "//div[@role='button' and @aria-label='Send']",
            ]

            send_button = None
            for xpath in possible_xpaths:
                try:
                    send_button = WebDriverWait(self.driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    if send_button:
                        # print(xpath)
                        break
                except Exception:
                    continue

            if not send_button:
                logging.warning(f"Send button not found for {contact['phone']}")
                print(f"Send button not found for {contact['phone']}")
                return False, None, "Send button not found"

            # Send the message
            self.cooldown(random.uniform(self.min_delay, self.max_delay), contact['name'])
            send_button.click()
            time.sleep(3)  # short pause to allow message to send

            logging.info(f"Message sent to {contact['phone']}")
            print(f"Message sent to {contact['phone']}")
            return True, message, None

        except Exception as e:
            logging.error(f"Error sending message to {contact.get('phone', '?')}: {str(e)}")
            print(f"Error sending message to {contact.get('phone', '?')}: {str(e)}")
            return False, None, "Error sending message"

    def process_numbers(self, messages, retry_mode=False, retry_contacts=None, progress_callback=None):
        """Process all contacts from Excel sheet"""
        try:
            if retry_mode and retry_contacts:
                contacts = retry_contacts
            else:
                contacts = self.load_data()

            # Random logic
            # If the checkbox is checked, shuffle the list before starting
            if self.config.random_order:
                random.shuffle(contacts)
                logging.info("Contacts shuffled for random order.")

            self.setup_driver()
            self.close_popups()

            results = {
                'successful': [],
                'failed': [],
                'messages_sent': {}  # Track which message was sent to each contact
            }

            for contact in contacts:
                if self.gui and self.gui.is_paused:
                    while self.gui.is_paused:
                        time.sleep(1)
                        if not self.gui.is_running:
                            return results

                # Randomly select a message variation
                message = random.choice(messages)
                success, sent_message, error = self.send_message(contact, message)

                if success:
                    results['successful'].append(contact['phone'])
                    results['messages_sent'][contact['phone']] = sent_message
                    if self.gui:
                        self.gui.update_status_in_tree(contact['phone'], 'Success ✓', 'Message sent')
                        self.gui.update_excel_status(contact['phone'], 'Success', 'Message sent', sent_message)
                else:
                    results['failed'].append({
                        'name': contact['name'],
                        'phone': contact['phone'],
                        'error': error
                    })
                    if self.gui:
                        self.gui.update_status_in_tree(contact['phone'], 'Failed ✗', f'Error sending message: {error}')
                        self.gui.update_excel_status(contact['phone'], 'Failed', f'Error sending message: {error}', None)

                if progress_callback:
                    progress_callback()

                # Random delay between messages (pause-aware)
                delay = random.uniform(5, 7)
                end_time = time.time() + delay
                while time.time() < end_time:
                    if self.gui and self.gui.is_paused:
                        time.sleep(0.5)
                        continue
                    time.sleep(0.2)  # small chunks so pause can interrupt

            return results

        except Exception as e:
            logging.error(f"Error in process_numbers: {str(e)}")
            raise
        finally:
            if self.driver:
                self.driver.quit()

if __name__ == "__main__":
    app = WhatsAppAutomationGUI()
    app.root.mainloop()
