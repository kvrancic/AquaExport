"""
Water Quality Data Exporter
===========================
A modern, high-performance tool for exporting water quality data from PostgreSQL to Excel.

Features:
- Beautiful GUI with date pickers
- Blazing fast performance using bulk operations
- Idempotent file handling (creates or updates yearly workbooks)
- Comprehensive error handling and logging
- Single-file executable support

Author: Water Quality Export System
Version: 2.0.0
"""

import os
import sys
import logging
import traceback
import asyncio
import shutil
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from collections import defaultdict
import tomli
import psycopg2
from psycopg2.extras import RealDictCursor
import openpyxl
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import threading


# Configure logging
def setup_logging(export_dir: Path) -> logging.Logger:
    """Set up rotating log file in export directory."""
    log_file = export_dir / "exporter.log"
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # File handler with rotation
    from logging.handlers import RotatingFileHandler
    file_handler = RotatingFileHandler(
        log_file, maxBytes=10*1024*1024, backupCount=5
    )
    file_handler.setFormatter(formatter)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    # Set up logger
    logger = logging.getLogger('WaterQualityExporter')
    logger.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger



@dataclass
class TagMapping:
    """Maps database tag indices to Excel columns for each location."""
    location: str
    mutnoca: Optional[int]
    klor: Optional[int] 
    temp: Optional[int]
    pH: Optional[int]
    redox: Optional[int]

@dataclass
class Config:
    """Application configuration."""
    db_host: str
    db_port: int
    db_name: str
    db_user: str
    db_password: str
    export_dir: Path
    template_path: Path
    
    @classmethod
    def from_file(cls, config_path: str = "config.toml") -> 'Config':
        """Load configuration from TOML file."""
        try:
            with open(config_path, "rb") as f:
                data = tomli.load(f)
            
            return cls(
                db_host=data["database"]["host"],
                db_port=data["database"]["port"],
                db_name=data["database"]["name"],
                db_user=data["database"]["user"],
                db_password=data["database"]["password"],
                export_dir=Path(data["export"]["directory"]),
                template_path=Path(data["export"]["template_path"])
            )
        except Exception as e:
            # Default configuration if file not found
            return cls(
                db_host="localhost",
                db_port=5432,
                db_name="SCADA_arhiva_rab",
                db_user="postgres",
                db_password="fornax123",
                export_dir=Path("./exports"),
                template_path=Path("./template.xlsx")
            )

class WaterQualityExporter:
    """Main exporter class handling database queries and Excel writing."""
    
    # Excel layout constants
    MONTH_NAMES = {
        1: "siječanj", 2: "veljača", 3: "ožujak", 4: "travanj",
        5: "svibanj", 6: "lipanj", 7: "srpanj", 8: "kolovoz",
        9: "rujan", 10: "listopad", 11: "studeni", 12: "prosinac"
    }
    
    BLOCK_ANCHORS = {
        'PK Barbat': 11,
        'VS Lopar': 59,
        'VS Perici': 107
    }
    
    def __init__(self, config: Config, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.conn = None
        
        # Load tag mappings from config if available, otherwise use defaults
        self.tag_mappings = self._load_tag_mappings()
        
    def _load_tag_mappings(self) -> Dict[str, TagMapping]:
        """Load tag mappings from config file or use defaults."""
        # Default mappings based on electrical engineers' PLC addresses
        default_mappings = {
            'PK Barbat': TagMapping(
                location='PK Barbat',
                mutnoca=3,      # N12:7 - pk_barb\mutnoca
                klor=21,        # N12:8 - pk_barb\trend_klora_izlaza
                temp=134,       # N12:10 - pk_barb\temperatura_vode
                pH=132,         # N12:9 - pk_barb\pH_vode
                redox=133       # N12:11 - pk_barb\redox
            ),
            'VS Lopar': TagMapping(
                location='VS Lopar',
                mutnoca=None,   # Not measured at this location
                klor=151,       # N22:3 - vsloparn\N22_3
                temp=155,       # N22:7 - vsloparn\N22_7
                pH=None,        # Not measured at this location
                redox=156       # N22:9 - vsloparn\N22_9
            ),
            'VS Perici': TagMapping(
                location='VS Perici',
                mutnoca=None,   # Not measured at this location
                klor=72,        # N15:40 - vs_perici\klor
                temp=82,        # N15:42 - vs_perici\temp_vode
                pH=None,        # Not measured at this location
                redox=81        # N15:41 - vs_perici\redox
            )
        }
        
        # Try to load from config file if tag_mappings section exists
        try:
            import tomli
            with open("config.toml", "rb") as f:
                config_data = tomli.load(f)
                
            if "tag_mappings" in config_data:
                self.logger.info("Loading tag mappings from config.toml")
                mappings = {}
                
                for location, tags in config_data["tag_mappings"].items():
                    # Convert location key to proper name
                    location_name = {
                        'pk_barbat': 'PK Barbat',
                        'vs_lopar': 'VS Lopar',
                        'vs_perici': 'VS Perici'
                    }.get(location.lower(), location)
                    
                    mappings[location_name] = TagMapping(
                        location=location_name,
                        mutnoca=tags.get('mutnoca'),
                        klor=tags.get('klor'),
                        temp=tags.get('temp'),
                        pH=tags.get('pH'),
                        redox=tags.get('redox')
                    )
                    
                return mappings
        except Exception as e:
            self.logger.debug(f"Using default tag mappings: {e}")
            
        return default_mappings
        
    def connect_db(self) -> None:
        """Establish database connection."""
        try:
            self.conn = psycopg2.connect(
                host=self.config.db_host,
                port=self.config.db_port,
                database=self.config.db_name,
                user=self.config.db_user,
                password=self.config.db_password,
                cursor_factory=RealDictCursor
            )
            self.logger.info("Database connection established")
        except Exception as e:
            self.logger.error(f"Database connection failed: {e}")
            raise
    
    def disconnect_db(self) -> None:
        """Close database connection."""
        if self.conn:
            self.conn.close()
            self.logger.info("Database connection closed")
    
    def fetch_data(self, start_date: date, end_date: date) -> Dict[str, Dict]:
        """
        Fetch aggregated data from database for date range.
        Returns nested dict: {location -> {date -> {param -> (min, max, avg)}}}
        """
        if not self.conn:
            self.connect_db()
            
        results = defaultdict(lambda: defaultdict(dict))
        
        try:
            with self.conn.cursor() as cur:
                # Query template for each parameter
                query = """
                    SELECT 
                        date_trunc('day', dateandtime) AS day,
                        COALESCE(MIN(CASE WHEN val > 0 THEN val END), 0) AS min_val,
                        COALESCE(MAX(val), 0) AS max_val,
                        COALESCE(AVG(CASE WHEN val > 0 THEN val END), 0) AS avg_val
                    FROM floattable
                    WHERE dateandtime >= %s 
                        AND dateandtime < %s + INTERVAL '1 day'
                        AND tagindex = %s
                    GROUP BY day
                    ORDER BY day
                """
                
                # Fetch data for each location and parameter
                for location, mapping in self.tag_mappings.items():
                    self.logger.info(f"Fetching data for {location}")
                    
                    params = {
                        'mutnoca': mapping.mutnoca,
                        'klor': mapping.klor,
                        'temp': mapping.temp,
                        'pH': mapping.pH,
                        'redox': mapping.redox
                    }
                    
                    for param_name, tag_index in params.items():
                        if tag_index is None:
                            continue
                            
                        cur.execute(query, (start_date, end_date, tag_index))
                        rows = cur.fetchall()
                        
                        for row in rows:
                            day = row['day'].date()
                            results[location][day][param_name] = (
                                round(row['min_val'], 2) if row['min_val'] else None,
                                round(row['max_val'], 2) if row['max_val'] else None,
                                round(row['avg_val'], 2) if row['avg_val'] else None
                            )
                        
                        self.logger.debug(f"  {param_name}: {len(rows)} days of data")
                
        except Exception as e:
            self.logger.error(f"Error fetching data: {e}")
            raise
            
        return dict(results)
    
    def get_or_create_workbook(self, year: int) -> Tuple[openpyxl.Workbook, Path]:
        """Get existing workbook or create new one from template."""
        # Ensure export directory exists
        self.config.export_dir.mkdir(parents=True, exist_ok=True)
        
        # Workbook path
        wb_path = self.config.export_dir / f"export_{year}.xlsx"
        
        if wb_path.exists():
            self.logger.info(f"Opening existing workbook: {wb_path}")
            try:
                wb = openpyxl.load_workbook(wb_path)
            except PermissionError:
                # File is likely open in Excel
                raise PermissionError(
                    f"Datoteka {wb_path.name} je trenutno otvorena u Excel-u!\n\n"
                    f"Molimo zatvorite datoteku u Excel-u i pokušajte ponovno.\n"
                    f"Datoteka: {wb_path}"
                )
        else:
            self.logger.info(f"Creating new workbook from template for year {year}")
            # Copy template
            shutil.copy2(self.config.template_path, wb_path)
            wb = openpyxl.load_workbook(wb_path)
            
            # Pre-fill year in all sheets
            for month in range(1, 13):
                sheet_name = f"P-{month:02d}"
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    # Fill year cells
                    for cell in ['B9', 'B57', 'B105']:
                        ws[cell] = year
                        
        return wb, wb_path
    
    def write_to_excel(self, data: Dict[str, Dict], start_date: date, end_date: date,
                      progress_callback=None) -> None:
        """Write data to Excel workbook."""
        # Group data by year
        years_data = defaultdict(lambda: defaultdict(dict))
        
        current = start_date
        while current <= end_date:
            for location, location_data in data.items():
                if current in location_data:
                    years_data[current.year][location][current] = location_data[current]
            current += timedelta(days=1)
        
        # Process each year
        total_days = (end_date - start_date).days + 1
        processed_days = 0
        
        for year, year_data in years_data.items():
            wb, wb_path = self.get_or_create_workbook(year)
            
            try:
                # Process each location's data
                for location, location_data in year_data.items():
                    block_anchor = self.BLOCK_ANCHORS.get(location)
                    if not block_anchor:
                        self.logger.warning(f"Unknown location: {location}")
                        continue
                    
                    for day_date, params in location_data.items():
                        month = day_date.month
                        day = day_date.day
                        
                        sheet_name = f"P-{month:02d}"
                        if sheet_name not in wb.sheetnames:
                            self.logger.warning(f"Sheet {sheet_name} not found")
                            continue
                            
                        ws = wb[sheet_name]
                        
                        # Calculate row (block_anchor + 2 + (day - 1))
                        row = block_anchor + 2 + (day - 1)
                        
                        # Write day number in column A
                        ws[f'A{row}'] = day
                        
                        # Write location name on first day of block
                        if day == 1:
                            ws[f'B{row}'] = location
                        
                        # Write parameter values - different column layouts per location
                        if location == 'PK Barbat':
                            # PK Barbat has all 5 parameters
                            col_mapping = {
                                'mutnoca': {'max': 'C', 'min': 'D', 'avg': 'E'},
                                'klor': {'max': 'F', 'min': 'G', 'avg': 'H'},
                                'temp': {'max': 'I', 'min': 'J', 'avg': 'K'},
                                'pH': {'max': 'L', 'min': 'M', 'avg': 'N'},
                                'redox': {'max': 'O', 'min': 'P', 'avg': 'Q'}
                            }
                        else:
                            # VS Lopar and VS Perici only have 3 parameters
                            col_mapping = {
                                'klor': {'max': 'C', 'min': 'D', 'avg': 'E'},
                                'temp': {'max': 'F', 'min': 'G', 'avg': 'H'},
                                'redox': {'max': 'I', 'min': 'J', 'avg': 'K'}
                            }
                        
                        for param, values in params.items():
                            if param in col_mapping and values:
                                min_val, max_val, avg_val = values
                                if max_val is not None:
                                    ws[f"{col_mapping[param]['max']}{row}"] = max_val
                                if min_val is not None:
                                    ws[f"{col_mapping[param]['min']}{row}"] = min_val
                                if avg_val is not None:
                                    ws[f"{col_mapping[param]['avg']}{row}"] = avg_val
                        
                        processed_days += 1
                        if progress_callback:
                            progress_callback(processed_days, total_days)
                
                # Save workbook
                try:
                    wb.save(wb_path)
                    self.logger.info(f"Saved workbook: {wb_path}")
                except PermissionError:
                    # File is likely open in Excel
                    raise PermissionError(
                        f"Datoteka {wb_path.name} je trenutno otvorena u Excel-u!\n\n"
                        f"Molimo zatvorite datoteku u Excel-u i pokušajte ponovno.\n"
                        f"Datoteka: {wb_path}"
                    )
                
            except Exception as e:
                self.logger.error(f"Error writing to Excel: {e}")
                raise
            finally:
                wb.close()

class ModernGUI:
    """Modern GUI for the Water Quality Exporter."""
    
    def __init__(self, config: Config, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.exporter = WaterQualityExporter(config, logger)
        
        self.root = tk.Tk()
        self.root.title("AquaExport Pro 2.0 - Izvoz Podataka Kvalitete Vode")
        self.root.geometry("600x720")
        self.root.resizable(False, False)
        
        # Set icon (if available)
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # Apply modern theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Custom colors
        self.bg_color = "#f0f4f8"
        self.primary_color = "#2196F3"
        self.success_color = "#4CAF50"
        self.error_color = "#f44336"
        
        self.root.configure(bg=self.bg_color)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create the user interface."""
        # Main container
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Logo and title
        self.add_header(main_frame)
        
        # Date selection frame
        self.add_date_selection(main_frame)
        
        # Export button
        self.add_export_button(main_frame)
        
        # Progress section
        self.add_progress_section(main_frame)
        
        # Status/log section
        self.add_status_section(main_frame)
        
        # Footer
        self.add_footer(main_frame)
        
        # Help button
        self.add_help_button(main_frame)
        
    def add_header(self, parent):
        """Add title header."""
        header_frame = tk.Frame(parent, bg=self.bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title
        title_frame = tk.Frame(header_frame, bg=self.bg_color)
        title_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = tk.Label(
            title_frame,
            text="AquaExport Pro 2.0",
            font=("Segoe UI", 24, "bold"),
            fg=self.primary_color,
            bg=self.bg_color
        )
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(
            title_frame,
            text="Izvoz podataka kvalitete vode",
            font=("Segoe UI", 12),
            fg="#666",
            bg=self.bg_color
        )
        subtitle_label.pack(anchor=tk.W)
        
    def add_date_selection(self, parent):
        """Add date selection widgets."""
        date_frame = tk.LabelFrame(
            parent,
            text="Odabir Datuma",
            font=("Segoe UI", 12, "bold"),
            bg=self.bg_color,
            fg=self.primary_color,
            relief=tk.FLAT,
            borderwidth=2
        )
        date_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Inner frame for padding
        inner_frame = tk.Frame(date_frame, bg=self.bg_color)
        inner_frame.pack(padx=20, pady=20)
        
        # Start date
        start_label = tk.Label(
            inner_frame,
            text="Početni datum:",
            font=("Segoe UI", 11),
            bg=self.bg_color
        )
        start_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))
        
        self.start_date = DateEntry(
            inner_frame,
            width=12,
            background=self.primary_color,
            foreground='white',
            borderwidth=2,
            font=("Segoe UI", 10),
            date_pattern='dd.mm.yyyy'
        )
        self.start_date.grid(row=0, column=1, padx=(0, 30), pady=(0, 10))
        
        # End date
        end_label = tk.Label(
            inner_frame,
            text="Završni datum:",
            font=("Segoe UI", 11),
            bg=self.bg_color
        )
        end_label.grid(row=0, column=2, sticky=tk.W, padx=(0, 10), pady=(0, 10))
        
        self.end_date = DateEntry(
            inner_frame,
            width=12,
            background=self.primary_color,
            foreground='white',
            borderwidth=2,
            font=("Segoe UI", 10),
            date_pattern='dd.mm.yyyy'
        )
        self.end_date.grid(row=0, column=3, pady=(0, 10))
        
        # Quick select buttons
        quick_frame = tk.Frame(inner_frame, bg=self.bg_color)
        quick_frame.grid(row=1, column=0, columnspan=4, pady=(10, 0))
        
        tk.Label(
            quick_frame,
            text="Brzi odabir:",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg="#666"
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        for text, days in [("Danas", 0), ("7 dana", 7), ("30 dana", 30), ("Godina", 365)]:
            btn = tk.Button(
                quick_frame,
                text=text,
                font=("Segoe UI", 9),
                bg="white",
                relief=tk.FLAT,
                borderwidth=1,
                padx=15,
                command=lambda d=days: self.set_date_range(d)
            )
            btn.pack(side=tk.LEFT, padx=2)
            
    def add_export_button(self, parent):
        """Add the main export button."""
        self.export_button = tk.Button(
            parent,
            text="IZVEZI PODATKE",
            font=("Segoe UI", 14, "bold"),
            bg=self.success_color,
            fg="white",
            relief=tk.FLAT,
            height=2,
            command=self.start_export
        )
        self.export_button.pack(fill=tk.X, pady=(0, 20))
        
    def add_progress_section(self, parent):
        """Add progress bar and status."""
        progress_frame = tk.Frame(parent, bg=self.bg_color)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_label = tk.Label(
            progress_frame,
            text="Spremno za izvoz",
            font=("Segoe UI", 10),
            bg=self.bg_color,
            fg="#666"
        )
        self.progress_label.pack(anchor=tk.W)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # Style the progress bar
        self.style.configure(
            "Custom.Horizontal.TProgressbar",
            background=self.primary_color,
            troughcolor="#e0e0e0",
            borderwidth=0,
            lightcolor=self.primary_color,
            darkcolor=self.primary_color
        )
        
    def add_status_section(self, parent):
        """Add status/log display."""
        status_frame = tk.LabelFrame(
            parent,
            text="Status",
            font=("Segoe UI", 10),
            bg=self.bg_color,
            relief=tk.FLAT,
            borderwidth=1
        )
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(
            status_frame,
            height=6,
            font=("Consolas", 9),
            bg="white",
            relief=tk.FLAT,
            borderwidth=1
        )
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.status_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)
        
    def add_footer(self, parent):
        """Add company footer."""
        footer_frame = tk.Frame(parent, bg=self.bg_color)
        footer_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Minimalistic company info
        footer_label = tk.Label(
            footer_frame,
            text="Created by FORNAX d.o.o • info@fornax-automatika.hr",
            font=("Segoe UI", 8),
            fg="#999",
            bg=self.bg_color
        )
        footer_label.pack(anchor=tk.CENTER)
            
    def add_help_button(self, parent):
        """Add help button."""
        help_button = tk.Button(
            parent,
            text="?",
            font=("Segoe UI", 12, "bold"),
            bg=self.primary_color,
            fg="white",
            relief=tk.FLAT,
            width=3,
            height=1,
            command=self.show_help
        )
        help_button.place(relx=1.0, rely=0, anchor=tk.NE)
        
    def set_date_range(self, days):
        """Set date range for quick selection."""
        end = datetime.now().date()
        if days == 0:
            start = end
        elif days == 365:
            start = date(end.year, 1, 1)
        else:
            start = end - timedelta(days=days)
            
        self.start_date.set_date(start)
        self.end_date.set_date(end)
        
    def log_status(self, message, level="INFO"):
        """Add message to status display."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()
        
    def update_progress(self, current, total):
        """Update progress bar."""
        if total > 0:
            progress = (current / total) * 100
            self.progress_bar['value'] = progress
            self.progress_label.config(text=f"Obrađeno: {current}/{total} dana")
            self.root.update()
            
    def start_export(self):
        """Start the export process in a separate thread."""
        self.export_button.config(state=tk.DISABLED, text="IZVOZ U TIJEKU...")
        self.progress_bar['value'] = 0
        self.status_text.delete(1.0, tk.END)
        
        # Run export in thread to keep GUI responsive
        thread = threading.Thread(target=self.run_export)
        thread.daemon = True
        thread.start()
        
    def run_export(self):
        """Run the actual export process."""
        try:
            start = self.start_date.get_date()
            end = self.end_date.get_date()
            
            if start > end:
                raise ValueError("Početni datum mora biti prije završnog datuma!")
            
            # Check for potentially open Excel files
            self.log_status("Provjera otvorenih Excel datoteka...")
            open_files = self.check_open_excel_files(start, end)
            if open_files:
                self.log_status(f"Upozorenje: {len(open_files)} datoteka možda je otvoreno", "WARNING")
                
            self.log_status(f"Početak izvoza: {start} - {end}")
            
            # Connect to database
            self.log_status("Povezivanje s bazom podataka...")
            self.exporter.connect_db()
            
            # Fetch data
            self.log_status("Dohvaćanje podataka...")
            data = self.exporter.fetch_data(start, end)
            
            if not data:
                self.log_status("Nema podataka za odabrani period!", "WARNING")
            else:
                total_records = sum(len(loc_data) for loc_data in data.values())
                self.log_status(f"Pronađeno {total_records} zapisa")
                
                # Write to Excel
                self.log_status("Pisanje u Excel...")
                self.exporter.write_to_excel(
                    data, start, end,
                    progress_callback=self.update_progress
                )
                
                self.log_status("Izvoz završen uspješno!", "SUCCESS")
                
                # Show success message and open file location
                self.root.after(
                    0,
                    lambda: self.show_success_and_open_folder()
                )
                
        except PermissionError as e:
            self.logger.error(f"Permission error during export: {e}")
            self.log_status(f"GREŠKA: Datoteka je otvorena u Excel-u!", "ERROR")
            
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Datoteka je otvorena",
                    f"Excel datoteka je trenutno otvorena!\n\n"
                    f"Molimo zatvorite datoteku u Excel-u i pokušajte ponovno.\n\n"
                    f"{str(e)}"
                )
            )
        except Exception as e:
            self.logger.error(f"Export failed: {e}\n{traceback.format_exc()}")
            self.log_status(f"GREŠKA: {str(e)}", "ERROR")
            
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Greška",
                    f"Izvoz nije uspio!\n\n{str(e)}\n\nDetalji su zapisani u log datoteku."
                )
            )
            
        finally:
            self.exporter.disconnect_db()
            self.root.after(0, self.reset_ui)
            
    def check_open_excel_files(self, start_date: date, end_date: date) -> List[str]:
        """Check if any Excel files that will be written to are potentially open."""
        open_files = []
        
        # Check each year that will be processed
        current = start_date
        while current <= end_date:
            year = current.year
            wb_path = self.config.export_dir / f"export_{year}.xlsx"
            
            if wb_path.exists():
                try:
                    # Try to open the file in write mode to check if it's locked
                    with open(wb_path, 'r+b') as f:
                        pass  # File is not locked
                except PermissionError:
                    open_files.append(wb_path.name)
            
            # Move to next year
            current = date(year + 1, 1, 1)
            if current > end_date:
                break
                
        return open_files
        
    def show_success_and_open_folder(self):
        """Show success message and open the export folder."""
        # Show success message
        result = messagebox.showinfo(
            "Uspjeh",
            f"Podaci su uspješno izvezeni!\n\nLokacija: {self.config.export_dir}\n\nKliknite OK da otvorite mapu s datotekama."
        )
        
        # Open the export folder in Windows Explorer
        if result == 'ok':
            try:
                os.startfile(str(self.config.export_dir))
            except Exception as e:
                self.logger.warning(f"Could not open folder: {e}")
                # Fallback: show the path in a message box
                messagebox.showinfo(
                    "Lokacija datoteka",
                    f"Datoteke su spremljene u:\n{self.config.export_dir}"
                )
    
    def reset_ui(self):
        """Reset UI after export."""
        self.export_button.config(state=tk.NORMAL, text="IZVEZI PODATKE")
        self.progress_label.config(text="Spremno za izvoz")
        
    def show_help(self):
        """Show help dialog."""
        help_text = """
AquaExport Pro 2.0 - Upute za korištenje

1. ODABIR DATUMA:
   • Odaberite početni i završni datum za izvoz
   • Koristite brze tipke za česte periode
   • Podaci se izvozе po danima

2. IZVOZ PODATAKA:
   • Kliknite "IZVEZI PODATKE" za početak
   • Pratite napredak u statusnoj traci
   • Izvoz može potrajati za velike periode

3. REZULTATI:
   • Excel datoteke se spremaju u mapu za izvoz
   • Svaka godina ima zasebnu datoteku
   • Postojeće datoteke se ažuriraju

4. NAPOMENE:
   • Ne zatvarajte program tijekom izvoza
   • Zatvorite Excel datoteke prije izvoza
   • Za probleme provjerite log datoteku
   • Kontakt: neven.vrancic@fornax-automatika.hr
        """
        
        messagebox.showinfo("Pomoć", help_text.strip())
        
    def run(self):
        """Start the GUI application."""
        self.root.mainloop()

def main():
    """Main entry point."""
    try:
        # Load configuration
        config = Config.from_file()
        
        # Ensure export directory exists
        config.export_dir.mkdir(parents=True, exist_ok=True)
        
        # Set up logging
        logger = setup_logging(config.export_dir)
        logger.info("Starting AquaExport Pro 2.0")
        
        # Create default config file if it doesn't exist
        if not Path("config.toml").exists():
            with open("config.toml", "w") as f:
                f.write('''[database]
host = "localhost"
port = 5432
name = "SCADA_arhiva_rab"
user = "postgres"
password = "fakepass"

[export]
directory = "./exports"
template_path = "./template.xlsx"
''')
            logger.info("Created default config.toml")
        
        # Check template exists
        if not config.template_path.exists():
            logger.error(f"Template not found: {config.template_path}")
            messagebox.showerror(
                "Greška",
                f"Excel predložak nije pronađen!\n\n{config.template_path}"
            )
            return
            
        # Start GUI
        app = ModernGUI(config, logger)
        app.run()
        
    except Exception as e:
        logger.error(f"Application error: {e}\n{traceback.format_exc()}")
        messagebox.showerror(
            "Kritična greška",
            f"Aplikacija se ne može pokrenuti!\n\n{str(e)}"
        )
        
if __name__ == "__main__":
    main()