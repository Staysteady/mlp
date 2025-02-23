"""
Script to manually capture midpoint data from Excel with stability checks.
Supports both Windows (win32com) and Mac (xlwings) for live Excel reading.
"""
from market_maker.data.models import Session, Snapshot, init_db
from market_maker.data.prompt_dates import calculate_days_between, get_prompt_date
from datetime import datetime, timedelta
import pandas as pd
import platform
import time
import os
import sys
import math
import re
import numpy as np

# Initialize database
engine = init_db()

# ANSI color codes
GREEN = '\033[32m'
RED = '\033[31m'
GRAY = '\033[90m'  # Dark gray for no data
RESET = '\033[0m'

def color_text(text, color):
    """Add color to text."""
    return f"{color}{text}{RESET}"

# Detect OS and import appropriate Excel interface
IS_WINDOWS = platform.system() == 'Windows'
if IS_WINDOWS:
    import win32com.client
else:
    import xlwings as xw

# Stability settings
STABILITY_DURATION = 4    # seconds price must remain stable
MIN_PRICE_CHANGE = 0.01  # minimum price change to log
POLL_INTERVAL = 0.5      # seconds between checks

def is_valid_spread(date1, date2, value):
    """
    Validate if a spread combination is valid.
    
    Args:
        date1 (str/datetime): First date in the spread
        date2 (str/datetime): Second date in the spread
        value (float): The spread value
        
    Returns:
        bool: True if spread is valid, False otherwise
    """
    # Skip if any of the inputs are None or empty
    if any(x is None for x in [date1, date2]):
        return False
        
    # Convert datetime objects to strings in MMM-YY format
    if isinstance(date1, datetime):
        date1 = date1.strftime("%b-%y").upper()
    if isinstance(date2, datetime):
        date2 = date2.strftime("%b-%y").upper()
        
    # Convert dates to strings if they aren't already
    date1 = str(date1) if date1 is not None else ''
    date2 = str(date2) if date2 is not None else ''
        
    # Filter out JAN-70 dates (placeholder for NaN/invalid)
    if 'JAN-70' in date1 or 'JAN-70' in date2:
        return False
        
    # Try to convert value to float and check if it's valid
    try:
        float_value = float(value)
        # Allow NaN values to pass through - we'll handle them separately
        if math.isnan(float_value):
            return True
    except (ValueError, TypeError):
        return False
        
    # Special cases for 'C' and '3M'
    special_dates = ['C', '3M']
    
    # If both dates are special cases, it's invalid
    if date1 in special_dates and date2 in special_dates:
        return False
        
    # If neither date is special, validate the format (MMM-YY)
    if date1 not in special_dates and not isinstance(date1, datetime):
        if not re.match(r'^[A-Z]{3}-\d{2}$', date1):
            return False
            
    if date2 not in special_dates and not isinstance(date2, datetime):
        if not re.match(r'^[A-Z]{3}-\d{2}$', date2):
            return False
            
    # Don't allow spreads between the same dates unless one is a special case
    if date1 == date2 and date1 not in special_dates:
        return False
        
    return True

class PricePoint:
    """Class to track price stability."""
    def __init__(self, spread, value, timestamp, bid=0.0, ask=0.0, bid_volume=0, ask_volume=0, is_primary=False):
        self.spread = spread
        self.value = self._safe_float_conversion(value)
        self.bid = self._safe_float_conversion(bid)
        self.ask = self._safe_float_conversion(ask)
        self.bid_volume = bid_volume
        self.ask_volume = ask_volume
        self.first_seen = timestamp
        self.last_seen = timestamp
        self.unchanged_since = timestamp
        self.is_stable = False  # Start as unstable
        self.is_recorded = False  # Track if we've recorded this value
        self.last_recorded_value = self._safe_float_conversion(value)  # Initialize with current value
        self.last_recorded_bid = self._safe_float_conversion(bid)  # Track last recorded bid
        self.last_recorded_ask = self._safe_float_conversion(ask)  # Track last recorded ask
        self.is_primary = is_primary  # Whether this is a primary spread (Section 1, rows 4-29)
        self.dependency = None  # Track which primary spread caused this change

    def _safe_float_conversion(self, value):
        """Safely convert a value to float, handling various types."""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            try:
                return float(value)
            except ValueError:
                return None
        return None  # Default for any other type (datetime, etc)
    
    def update(self, value, timestamp, bid=None, ask=None, bid_volume=None, ask_volume=None, dependency=None):
        """Update price point with new values."""
        try:
            current_value = self._safe_float_conversion(value)
            current_bid = self._safe_float_conversion(bid) if bid is not None else self.bid
            current_ask = self._safe_float_conversion(ask) if ask is not None else self.ask
            
            # Update volumes without tracking changes
            if bid_volume is not None:
                self.bid_volume = bid_volume
            if ask_volume is not None:
                self.ask_volume = ask_volume
            
            # Check if any price has changed significantly
            mid_changed = False
            bid_changed = False
            ask_changed = False
            
            if current_value is not None and self.value is not None:
                mid_changed = abs(self.value - current_value) >= MIN_PRICE_CHANGE
            elif current_value != self.value:  # One is None and the other isn't
                mid_changed = True
                
            if current_bid is not None and self.bid is not None:
                bid_changed = abs(self.bid - current_bid) >= MIN_PRICE_CHANGE
            elif current_bid != self.bid:  # One is None and the other isn't
                bid_changed = True
                
            if current_ask is not None and self.ask is not None:
                ask_changed = abs(self.ask - current_ask) >= MIN_PRICE_CHANGE
            elif current_ask != self.ask:  # One is None and the other isn't
                ask_changed = True
            
            if mid_changed or bid_changed or ask_changed:
                # Price changed significantly, reset stability tracking
                old_value = self.last_recorded_value
                old_bid = self.last_recorded_bid
                old_ask = self.last_recorded_ask
                
                self.value = current_value
                self.bid = current_bid
                self.ask = current_ask
                self.unchanged_since = timestamp
                self.is_stable = False
                self.is_recorded = False  # Reset recorded flag on significant change
                if dependency:
                    self.dependency = dependency  # Update dependency if provided
                return False, (old_value, old_bid, old_ask)
            else:
                # Price hasn't changed significantly
                self.last_seen = timestamp
                # Check if price has been stable for required duration
                stability_duration = (timestamp - self.unchanged_since).total_seconds()
                was_stable = self.is_stable
                self.is_stable = stability_duration >= STABILITY_DURATION
                # Return True only when we first become stable and haven't recorded yet
                return (self.is_stable and not was_stable and not self.is_recorded), (self.last_recorded_value, self.last_recorded_bid, self.last_recorded_ask)
        except Exception as e:
            print(f"Error updating price point: {e}")
            return False, (self.last_recorded_value, self.last_recorded_bid, self.last_recorded_ask)

    def mark_recorded(self):
        """Mark this price point as having been recorded."""
        self.is_recorded = True
        self.last_recorded_value = self.value  # Update last recorded value
        self.last_recorded_bid = self.bid  # Update last recorded bid
        self.last_recorded_ask = self.ask  # Update last recorded ask

def format_date(date_val):
    """Format date value to MMM-YY format."""
    if pd.isna(date_val):
        return None
    try:
        # Handle special cases
        if isinstance(date_val, str):
            if date_val == "3M":  # Special case in your data
                return "3M"
            if date_val == "C":   # Special case for current
                return "C"
            if len(date_val) <= 7:  # Already in correct format
                return date_val.upper()
            # Try to parse string date
            date_val = pd.to_datetime(date_val)
        
        # Convert to datetime if needed
        if not isinstance(date_val, pd.Timestamp):
            date_val = pd.to_datetime(date_val)
            
        # Format to MMM-YY
        return date_val.strftime("%b-%y").upper()
            
    except Exception as e:
        print(f"Warning: Could not format date {date_val}: {e}")
        return None

class ExcelInterface:
    """Abstract base class for Excel interfaces."""
    def __init__(self):
        if IS_WINDOWS:
            self.setup_windows()
        else:
            self.setup_mac()
    
    def setup_windows(self):
        """Setup Windows COM interface."""
        try:
            self.excel = win32com.client.GetObject(None, "Excel.Application")
            self.wb = None
            # Find our workbook in open workbooks
            for wb in self.excel.Workbooks:
                if wb.Name == "NEON_ML.xlsm":
                    self.wb = wb
                    break
            if not self.wb:
                raise Exception("NEON_ML.xlsm not found in open workbooks")
            self.sheet = self.wb.Worksheets("AH NEON")
            self.sod_sheet = self.wb.Worksheets("SOD")
        except Exception as e:
            print(f"Error connecting to Excel on Windows: {e}")
            raise
    
    def setup_mac(self):
        """Setup Mac xlwings interface."""
        try:
            # First try to get existing Excel instance
            apps = xw.apps
            if len(apps) > 0:
                app = apps.active
            else:
                # If no Excel instance, start a new one
                app = xw.App()
            
            # Try to get existing workbook
            try:
                self.wb = app.books['NEON_ML.xlsm']
            except:
                # If workbook not found, try to open it from config path
                from market_maker.config.settings import EXCEL_FILE
                self.wb = app.books.open(EXCEL_FILE)
            
            self.sheet = self.wb.sheets["AH NEON"]
            self.sod_sheet = self.wb.sheets["SOD"]
            
            # Keep reference to app
            self.app = app
            
        except Exception as e:
            print(f"\nError connecting to Excel on Mac: {e}")
            print("Please ensure:")
            print("1. Excel is running")
            print("2. NEON_ML.xlsm is open")
            print("3. You have necessary permissions")
            raise
    
    def read_range(self, sheet, start_cell, nrows, ncols):
        """Read a range of values from Excel sheet."""
        try:
            # Parse start cell (e.g., 'A1' -> row 1, col 'A')
            start_col = start_cell[0].upper()
            if len(start_cell) > 2 and start_cell[1].isalpha():  # Handle 'AA4' style references
                start_col = start_cell[:2].upper()
                start_row = int(start_cell[2:])
            else:
                start_row = int(start_cell[1:])
            
            if IS_WINDOWS:
                # Windows COM range reading
                end_col = chr(ord(start_col[-1]) + ncols - 1)
                end_row = start_row + nrows - 1
                range_address = f"{start_cell}:{end_col}{end_row}"
                values = sheet.Range(range_address).Value
                if isinstance(values, tuple):
                    values = [list(row) for row in values]
                elif not isinstance(values, list):
                    values = [[values]]
            else:
                # xlwings range reading - read row by row for more reliability
                values = []
                for row in range(nrows):
                    row_values = []
                    for col in range(ncols):
                        # Calculate column letter
                        if len(start_col) == 1:
                            col_letter = chr(ord(start_col) + col)
                        else:
                            # Handle double letter columns (AA, AB, etc)
                            base = ord('Z') - ord('A') + 1
                            first = chr(ord('A') + ((ord(start_col[0]) - ord('A')) * base + col) // base)
                            second = chr(ord('A') + ((ord(start_col[0]) - ord('A')) * base + col) % base)
                            col_letter = f"{first}{second}"
                        
                        cell_ref = f"{col_letter}{start_row + row}"
                        try:
                            value = sheet.range(cell_ref).value
                            # Convert numeric values to float, handle None
                            if value is None:
                                value = 0.0
                            elif isinstance(value, (int, float)):
                                value = float(value)
                            row_values.append(value)
                        except Exception as e:
                            print(f"Warning: Error reading cell {cell_ref}: {e}")
                            row_values.append(0.0)
                    values.append(row_values)
            
            return values
            
        except Exception as e:
            print(f"Error reading range starting at {start_cell}: {e}")
            # Return empty list with correct dimensions
            return [[0.0] * ncols for _ in range(nrows)]
    
    def read_cell(self, sheet, cell):
        """Read a single cell value."""
        try:
            if IS_WINDOWS:
                value = sheet.Range(cell).Value
            else:
                value = sheet.range(cell).value
            
            # Convert numeric values to float
            if isinstance(value, (int, float)):
                return float(value)
            return value
        except Exception as e:
            print(f"Error reading cell {cell}: {e}")
            return None

class ExcelMonitor:
    """Class to maintain Excel connection and track prices."""
    def __init__(self):
        self.price_tracker = {}
        self.session = Session()
        self.excel = None
        self.c_date = None  # Store C (cash) date
        self.three_m_date = None  # Store 3M date
        self.spread_prefix = None  # Store spread prefix from B2
        self.connect_to_excel()
        self.update_reference_dates()
        self.load_spread_prefix()
    
    def load_spread_prefix(self):
        """Load the spread prefix from cell B2."""
        if not self.ensure_excel_connection():
            return
            
        try:
            self.spread_prefix = self.excel.read_cell(self.excel.sheet, "B2")
            if not self.spread_prefix:
                print("\nWarning: Could not read spread prefix from B2, using default")
                self.spread_prefix = "AHD"  # Default prefix if not found
        except Exception as e:
            print(f"\nError reading spread prefix: {e}")
            self.spread_prefix = "AHD"  # Default prefix if error
    
    def format_spread_name(self, date1, date2):
        """Format spread name using prefix and without hyphens in month codes."""
        # Handle special cases (C and 3M)
        if date1 in ['C', '3M']:
            # For cash or 3M first leg, format as AHDCASH-FEB25 or AHD3M-FEB25
            leg1 = date1
            # Remove hyphen from second leg (e.g., FEB-25 -> FEB25)
            leg2 = date2.replace('-', '') if isinstance(date2, str) else date2.strftime("%b%y").upper()
            return f"{self.spread_prefix}{leg1}-{leg2}"
        elif date2 in ['C', '3M']:
            # Remove hyphen from first leg
            leg1 = date1.replace('-', '') if isinstance(date1, str) else date1.strftime("%b%y").upper()
            leg2 = date2
            return f"{self.spread_prefix}{leg1}-{leg2}"
        else:
            # Regular spread, remove hyphens from both legs
            leg1 = date1.replace('-', '') if isinstance(date1, str) else date1.strftime("%b%y").upper()
            leg2 = date2.replace('-', '') if isinstance(date2, str) else date2.strftime("%b%y").upper()
            return f"{self.spread_prefix}{leg1}-{leg2}"
    
    def update_reference_dates(self):
        """Update C and 3M reference dates from SOD sheet."""
        if not self.ensure_excel_connection():
            return
            
        try:
            # Get reference dates from SOD sheet
            self.c_date = self.excel.read_cell(self.excel.sod_sheet, "C8")
            self.three_m_date = self.excel.read_cell(self.excel.sod_sheet, "C9")
            
            if self.c_date is None or self.three_m_date is None:
                print("\nWarning: Could not read reference dates from SOD sheet")
        except Exception as e:
            print(f"\nError reading reference dates: {e}")
    
    def connect_to_excel(self):
        """Establish connection to Excel."""
        try:
            self.excel = ExcelInterface()
            return True
        except Exception as e:
            print(f"\nError connecting to Excel: {e}")
            print("Please ensure:")
            print("1. Excel is running")
            print("2. NEON_ML.xlsm is open")
            print("3. You have necessary permissions")
            return False
    
    def ensure_excel_connection(self):
        """Ensure we have a valid Excel connection."""
        if self.excel is None:
            return self.connect_to_excel()
        return True
    
    def read_excel_data(self):
        """Read all required data from Excel."""
        if not self.ensure_excel_connection():
            return None, None, None
            
        try:
            # Get reference dates
            c_date = self.excel.read_cell(self.excel.sod_sheet, "C8")
            three_m_date = self.excel.read_cell(self.excel.sod_sheet, "C9")
            
            # If we can't read basic cells, Excel connection might be broken
            if c_date is None or three_m_date is None:
                print("\nError: Cannot read reference dates from Excel")
                self.excel = None  # Force reconnection next time
                return None, None, None
            
            # Format dates for display
            c_display = c_date.strftime("%d/%m/%Y") if isinstance(c_date, datetime) else str(c_date)
            three_m_display = three_m_date.strftime("%d/%m/%Y") if isinstance(three_m_date, datetime) else str(three_m_date)
            
            # Read primary section (A-C) column by column
            primary_dates1 = [row[0] for row in self.excel.read_range(self.excel.sheet, "A4", 64, 1)]
            primary_dates2 = [row[0] for row in self.excel.read_range(self.excel.sheet, "B4", 64, 1)]
            primary_mids = [row[0] for row in self.excel.read_range(self.excel.sheet, "C4", 64, 1)]
            
            # Read derived1 section (Z-AB) column by column
            derived1_dates1 = [row[0] for row in self.excel.read_range(self.excel.sheet, "Z4", 83, 1)]
            derived1_dates2 = [row[0] for row in self.excel.read_range(self.excel.sheet, "AA4", 83, 1)]
            derived1_mids = [row[0] for row in self.excel.read_range(self.excel.sheet, "AB4", 83, 1)]
            
            # Read derived2 section (AW-AY) column by column
            derived2_dates1 = [row[0] for row in self.excel.read_range(self.excel.sheet, "AW4", 81, 1)]
            derived2_dates2 = [row[0] for row in self.excel.read_range(self.excel.sheet, "AX4", 81, 1)]
            derived2_mids = [row[0] for row in self.excel.read_range(self.excel.sheet, "AY4", 81, 1)]
            
            # Combine into DataFrames
            data = {
                'primary_dates': pd.DataFrame({
                    'prompt1': primary_dates1,
                    'prompt2': primary_dates2
                }),
                'derived1_dates': pd.DataFrame({
                    'prompt1': derived1_dates1,
                    'prompt2': derived1_dates2
                }),
                'derived2_dates': pd.DataFrame({
                    'prompt1': derived2_dates1,
                    'prompt2': derived2_dates2
                }),
                'primary_mids': pd.Series(primary_mids),
                'derived1_mids': pd.Series(derived1_mids),
                'derived2_mids': pd.Series(derived2_mids)
            }
            
            return data, (c_display, c_date), (three_m_display, three_m_date)
            
        except Exception as e:
            print(f"\nError reading Excel data: {e}")
            self.excel = None  # Force reconnection next time
            return None, None, None

    def capture_midpoints(self):
        """Capture current midpoint values from Excel."""
        try:
            # Get the current values
            current_values = {}
            seen_spreads = set()  # Track spreads we've seen to avoid duplicates
            all_changes = []  # Track all changes for unified sorting and printing
            capture_time = datetime.utcnow()
            
            # Section 1 (A-C)
            for row in range(2, 100):  # Adjust range as needed
                date1 = self.excel.read_cell(self.excel.sheet, f"A{row}")
                date2 = self.excel.read_cell(self.excel.sheet, f"B{row}")
                mid = self.excel.read_cell(self.excel.sheet, f"C{row}")
                
                self._process_spread_data(date1, date2, mid, seen_spreads, all_changes, capture_time, current_values, row, section=0)

            # Section 2 (Z-AB)
            for row in range(2, 100):
                date1 = self.excel.read_cell(self.excel.sheet, f"Z{row}")
                date2 = self.excel.read_cell(self.excel.sheet, f"AA{row}")
                mid = self.excel.read_cell(self.excel.sheet, f"AB{row}")
                
                self._process_spread_data(date1, date2, mid, seen_spreads, all_changes, capture_time, current_values, row, section=1)

            # Section 3 (AW-AY)
            for row in range(2, 100):
                date1 = self.excel.read_cell(self.excel.sheet, f"AW{row}")
                date2 = self.excel.read_cell(self.excel.sheet, f"AX{row}")
                mid = self.excel.read_cell(self.excel.sheet, f"AY{row}")
                
                self._process_spread_data(date1, date2, mid, seen_spreads, all_changes, capture_time, current_values, row, section=2)
            
            # If we have changes, commit them and display in a unified, sorted section
            if all_changes:
                self.session.commit()
                
                # Sort all changes by days first, then spread name
                def get_days(x):
                    if len(x) < 8 or x[7] is None:
                        return float('inf')
                    return x[7]
                
                # Sort strictly by days first, then spread name
                all_changes.sort(key=lambda x: (get_days(x), x[2]))
                
                # Print header only once
                print("\nTime                      | Type | A/D | Spread          | Days |"
                      " Mid Old  | Mid New  |  Mid Δ  |"
                      " Bid Old  | Bid New  |  Bid Δ  |"
                      " Ask Old  | Ask New  |  Ask Δ")
                print("-" * 180)
                
                # Print all changes
                for change_data in all_changes:
                    time, type_, spread, new_value, old_value, dependency, days, spread_type, new_bid, old_bid, new_ask, old_ask = change_data
                    
                    days_str = str(days) if days is not None else "N/A"
                    
                    # Calculate changes
                    mid_change = new_value - old_value
                    bid_change = new_bid - old_bid if not math.isnan(new_bid) and not math.isnan(old_bid) else None
                    ask_change = new_ask - old_ask if not math.isnan(new_ask) and not math.isnan(old_ask) else None
                    
                    # Format bid values
                    if math.isnan(new_bid) or math.isnan(old_bid) or (new_bid == 0 and old_bid == 0):
                        bid_old_str = color_text("  N/A  ", GRAY)
                        bid_new_str = color_text("  N/A  ", GRAY)
                        bid_change_str = color_text(" N/A ", GRAY)
                    else:
                        bid_old_str = f"{old_bid:^8.2f}"
                        bid_new_str = f"{new_bid:^8.2f}"
                        bid_change_str = f"{bid_change:^+7.2f}"
                        if bid_change > 0:
                            bid_change_str = color_text(bid_change_str, GREEN)
                        elif bid_change < 0:
                            bid_change_str = color_text(bid_change_str, RED)
                    
                    # Format ask values
                    if math.isnan(new_ask) or math.isnan(old_ask) or (new_ask == 0 and old_ask == 0):
                        ask_old_str = color_text("  N/A  ", GRAY)
                        ask_new_str = color_text("  N/A  ", GRAY)
                        ask_change_str = color_text(" N/A ", GRAY)
                    else:
                        ask_old_str = f"{old_ask:^8.2f}"
                        ask_new_str = f"{new_ask:^8.2f}"
                        ask_change_str = f"{ask_change:^+7.2f}"
                        if ask_change > 0:
                            ask_change_str = color_text(ask_change_str, GREEN)
                        elif ask_change < 0:
                            ask_change_str = color_text(ask_change_str, RED)
                    
                    # Format mid change
                    mid_change_str = f"{mid_change:^+7.2f}"
                    if mid_change > 0:
                        mid_change_str = color_text(mid_change_str, GREEN)
                    elif mid_change < 0:
                        mid_change_str = color_text(mid_change_str, RED)
                    
                    print(f"{time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]:<25} |"
                          f" {type_:^4} |"
                          f" {spread_type:^3} |"
                          f" {spread:<15} |"
                          f" {days_str:>4} |"
                          f" {old_value:^8.2f} |"
                          f" {new_value:^8.2f} |"
                          f" {mid_change_str} |"
                          f" {bid_old_str} |"
                          f" {bid_new_str} |"
                          f" {bid_change_str} |"
                          f" {ask_old_str} |"
                          f" {ask_new_str} |"
                          f" {ask_change_str}")
            
            return current_values
        except Exception as e:
            print(f"Error capturing midpoints: {e}")
            self.session.rollback()  # Rollback on error
            return {}

    def _process_spread_data(self, date1, date2, mid, seen_spreads, all_changes, capture_time, current_values, row, section=0):
        """Helper method to process spread data from any section."""
        # Format dates if they are datetime objects
        if isinstance(date1, datetime):
            date1 = date1.strftime("%b-%y").upper()
        if isinstance(date2, datetime):
            date2 = date2.strftime("%b-%y").upper()
        
        # Skip if we've already seen this spread combination
        spread_key = f"{date1}-{date2}"
        if spread_key in seen_spreads:
            return
        
        if is_valid_spread(date1, date2, mid):
            try:
                current_value = float(mid) if not isinstance(mid, datetime) else 0.0
                spread = self.format_spread_name(date1, date2)  # Use new formatting method
                
                # Read Fidessa data based on section
                bid_volume = 0
                bid_price = 0.0
                ask_price = 0.0
                ask_volume = 0
                
                try:
                    if section == 0:
                        # Section 1: E, F, G, H
                        bid_volume = self.excel.read_cell(self.excel.sheet, f"E{row}")
                        bid_price = self.excel.read_cell(self.excel.sheet, f"F{row}")
                        ask_price = self.excel.read_cell(self.excel.sheet, f"G{row}")
                        ask_volume = self.excel.read_cell(self.excel.sheet, f"H{row}")
                    elif section == 1:
                        # Section 2: AD, AE, AF, AG
                        bid_volume = self.excel.read_cell(self.excel.sheet, f"AD{row}")
                        bid_price = self.excel.read_cell(self.excel.sheet, f"AE{row}")
                        ask_price = self.excel.read_cell(self.excel.sheet, f"AF{row}")
                        ask_volume = self.excel.read_cell(self.excel.sheet, f"AG{row}")
                    else:
                        # Section 3: BA, BB, BC, BD
                        bid_volume = self.excel.read_cell(self.excel.sheet, f"BA{row}")
                        bid_price = self.excel.read_cell(self.excel.sheet, f"BB{row}")
                        ask_price = self.excel.read_cell(self.excel.sheet, f"BC{row}")
                        ask_volume = self.excel.read_cell(self.excel.sheet, f"BD{row}")
                    
                    # Convert to proper types
                    bid_volume = int(bid_volume) if bid_volume is not None else 0
                    ask_volume = int(ask_volume) if ask_volume is not None else 0
                    bid_price = float(bid_price) if bid_price is not None else None
                    ask_price = float(ask_price) if ask_price is not None else None
                    
                except Exception as e:
                    # Just log the error but continue processing
                    print(f"Warning: Could not read Fidessa data for {spread}: {e}")
                
                # Calculate days between for the spread
                days = None
                if date1 == 'C' and self.c_date:
                    date1_obj = self.c_date
                    date2_obj = get_prompt_date(date2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                elif date1 == '3M' and self.three_m_date:
                    date1_obj = self.three_m_date
                    date2_obj = get_prompt_date(date2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                elif date2 == '3M' and self.three_m_date:
                    date1_obj = get_prompt_date(date1)
                    date2_obj = self.three_m_date
                    if date1_obj:
                        days = abs((date2_obj - date1_obj).days)
                else:
                    days = calculate_days_between(date1, date2)
                
                # Determine if this is a primary spread (Section 1, rows 4-29)
                is_primary = section == 0 and 4 <= row <= 29
                primary_dependency = None
                spread_type = "A" if is_primary else "D"
                
                # Check if any primary spreads have changed recently
                for key, point in self.price_tracker.items():
                    if point.is_primary and not point.is_recorded and point.unchanged_since >= capture_time - timedelta(seconds=POLL_INTERVAL):
                        primary_dependency = key
                        break
                
                # Check if this is a new spread or value has changed
                if spread not in self.price_tracker:
                    # New spread - start tracking but don't show in changes
                    self.price_tracker[spread] = PricePoint(spread, current_value, capture_time, 
                                                          bid=bid_price, ask=ask_price,
                                                          bid_volume=bid_volume, ask_volume=ask_volume,
                                                          is_primary=is_primary)
                    if current_value != 0:  # Only record non-zero values
                        snapshot = Snapshot(
                            timestamp=capture_time,
                            spread_name=spread,
                            prompt1=date1,
                            prompt2=date2,
                            old_midpoint=current_value,
                            new_midpoint=current_value,
                            old_bid=0.0 if bid_price is None else bid_price,
                            new_bid=0.0 if bid_price is None else bid_price,
                            old_ask=0.0 if ask_price is None else ask_price,
                            new_ask=0.0 if ask_price is None else ask_price
                        )
                        self.session.add(snapshot)
                        self.price_tracker[spread].mark_recorded()
                else:
                    price_point = self.price_tracker[spread]
                    should_record, (old_value, old_bid, old_ask) = price_point.update(
                        current_value, capture_time,
                        bid=bid_price, ask=ask_price,
                        bid_volume=bid_volume, ask_volume=ask_volume,
                        dependency=primary_dependency
                    )
                    
                    # Check if value has changed significantly
                    mid_changed = abs(current_value - price_point.last_recorded_value) >= MIN_PRICE_CHANGE
                    bid_changed = abs(bid_price - price_point.last_recorded_bid) >= MIN_PRICE_CHANGE if bid_price is not None and price_point.last_recorded_bid is not None else False
                    ask_changed = abs(ask_price - price_point.last_recorded_ask) >= MIN_PRICE_CHANGE if ask_price is not None and price_point.last_recorded_ask is not None else False
                    
                    if mid_changed or bid_changed or ask_changed:
                        # Value has changed significantly
                        if current_value != 0:  # Only show non-zero values
                            # Record the change after stability period
                            if should_record:
                                # Add dependency info to changes list if this is a derived spread
                                if not price_point.is_primary and price_point.dependency:
                                    all_changes.append((capture_time, "CHG", spread, current_value, price_point.last_recorded_value, 
                                                      price_point.dependency, days, spread_type,
                                                      bid_price if bid_price is not None else np.nan, price_point.last_recorded_bid if price_point.last_recorded_bid is not None else np.nan,  # Use last recorded bid
                                                      ask_price if ask_price is not None else np.nan, price_point.last_recorded_ask if price_point.last_recorded_ask is not None else np.nan)) # Use last recorded ask
                                else:
                                    all_changes.append((capture_time, "CHG", spread, current_value, price_point.last_recorded_value, 
                                                      None, days, spread_type,
                                                      bid_price if bid_price is not None else np.nan, price_point.last_recorded_bid if price_point.last_recorded_bid is not None else np.nan,  # Use last recorded bid
                                                      ask_price if ask_price is not None else np.nan, price_point.last_recorded_ask if price_point.last_recorded_ask is not None else np.nan)) # Use last recorded ask
                                
                                snapshot = Snapshot(
                                    timestamp=capture_time,
                                    spread_name=spread,
                                    prompt1=date1,
                                    prompt2=date2,
                                    old_midpoint=price_point.last_recorded_value,
                                    new_midpoint=current_value,
                                    old_bid=0.0 if price_point.last_recorded_bid is None else price_point.last_recorded_bid,
                                    new_bid=0.0 if bid_price is None else bid_price,
                                    old_ask=0.0 if price_point.last_recorded_ask is None else price_point.last_recorded_ask,
                                    new_ask=0.0 if ask_price is None else ask_price
                                )
                                self.session.add(snapshot)
                                price_point.mark_recorded()
                
                current_values[spread] = current_value
                seen_spreads.add(spread_key)
            except Exception as e:
                # Only print warning for truly invalid spreads
                if not isinstance(e, (ValueError, TypeError)):
                    print(f"Error processing spread {date1}-{date2}: {e}")

def clear_screen():
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_full_snapshot(monitor, capture_time):
    """Print a full snapshot of all sections."""
    print("\nInitial Market Snapshot at", capture_time.strftime("%Y-%m-%d %H:%M:%S"))
    print("=" * 120)  # Widened to accommodate new columns
    
    # Print reference dates if available
    if monitor.c_date and monitor.three_m_date:
        print(f"\nReference Dates:")
        print(f"C (Cash): {monitor.c_date.strftime('%Y-%m-%d') if isinstance(monitor.c_date, datetime) else monitor.c_date}")
        print(f"3M:      {monitor.three_m_date.strftime('%Y-%m-%d') if isinstance(monitor.three_m_date, datetime) else monitor.three_m_date}")
    
    sections = [
        ("Section 1", "A", "B", "C", "E", "F", "G", "H"),  # Added Fidessa columns
        ("Section 2", "Z", "AA", "AB", "AD", "AE", "AF", "AG"),  # Added Fidessa columns
        ("Section 3", "AW", "AX", "AY", "BA", "BB", "BC", "BD")  # Added Fidessa columns
    ]
    
    # Initialize price points for monitoring without waiting for stability
    for section_idx, (section_name, col1, col2, col3, bid_vol_col, bid_col, ask_col, ask_vol_col) in enumerate(sections):
        print(f"\n{section_name}")
        print("-" * 140)  # Increased width
        print(f"{'Spread':<15} | {'A/D':^3} | {'Midpoint':>12} | {'Bid':>12} | {'Ask':>12} | {'Days Between':>12} | {'Dates':>35}")
        print("-" * 140)  # Increased width
        
        for row in range(4, 100):  # Start from row 4 where data begins
            date1 = monitor.excel.read_cell(monitor.excel.sheet, f"{col1}{row}")
            date2 = monitor.excel.read_cell(monitor.excel.sheet, f"{col2}{row}")
            mid = monitor.excel.read_cell(monitor.excel.sheet, f"{col3}{row}")
            
            # Read Fidessa data
            bid_volume = monitor.excel.read_cell(monitor.excel.sheet, f"{bid_vol_col}{row}")
            bid_price = monitor.excel.read_cell(monitor.excel.sheet, f"{bid_col}{row}")
            ask_price = monitor.excel.read_cell(monitor.excel.sheet, f"{ask_col}{row}")
            ask_volume = monitor.excel.read_cell(monitor.excel.sheet, f"{ask_vol_col}{row}")
            
            if is_valid_spread(date1, date2, mid):
                # Format spread name using new format
                spread = monitor.format_spread_name(date1, date2)
                
                # Determine if this is an actual spread (only in Section 1, rows 4-29)
                spread_type = "A" if section_idx == 0 and 4 <= row <= 29 else "D"
                
                # Calculate days between prompts
                days = None
                dates_str = ""
                
                # Handle special cases with C and 3M
                if date1 == 'C' and monitor.c_date:
                    date1_obj = monitor.c_date
                    date2_obj = get_prompt_date(date2)  # This will now return third Wednesday
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                        dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                elif date1 == '3M' and monitor.three_m_date:
                    date1_obj = monitor.three_m_date
                    date2_obj = get_prompt_date(date2)  # This will now return third Wednesday
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                        dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                elif date2 == '3M' and monitor.three_m_date:
                    date1_obj = get_prompt_date(date1)  # This will now return third Wednesday
                    date2_obj = monitor.three_m_date
                    if date1_obj:
                        days = abs((date2_obj - date1_obj).days)
                        dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                else:
                    days = calculate_days_between(date1, date2)  # This already uses third Wednesdays
                    if days is not None:
                        date1_obj = get_prompt_date(date1)  # This will now return third Wednesday
                        date2_obj = get_prompt_date(date2)  # This will now return third Wednesday
                        if date1_obj and date2_obj:
                            dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                
                days_str = str(days) if days is not None else "N/A"
                
                # Convert bid/ask to float and handle None values
                bid_price = float(bid_price) if bid_price is not None else None
                ask_price = float(ask_price) if ask_price is not None else None
                
                # Format bid/ask display - show empty space when no volume
                bid_display = f"{bid_price:12.2f}" if bid_price is not None else " " * 12
                ask_display = f"{ask_price:12.2f}" if ask_price is not None else " " * 12
                
                print(f"{spread:<15} | {spread_type:^3} | {float(mid):12.2f} | {bid_display} | {ask_display} | {days_str:>12} | {dates_str:>35}")
                
                # Initialize price point for monitoring but mark as recorded
                if spread not in monitor.price_tracker:
                    price_point = PricePoint(spread, mid, capture_time, 
                                          bid=bid_price, ask=ask_price,
                                          bid_volume=bid_volume, ask_volume=ask_volume,
                                          is_primary=spread_type == "A")
                    price_point.mark_recorded()  # Mark as recorded so we don't show it again immediately
                    monitor.price_tracker[spread] = price_point
                    
                    # Record initial snapshot in database
                    snapshot = Snapshot(
                        timestamp=capture_time,
                        spread_name=spread,
                        prompt1=date1,
                        prompt2=date2,
                        old_midpoint=float(mid),
                        new_midpoint=float(mid),
                        old_bid=0.0 if bid_price is None else bid_price,
                        new_bid=0.0 if bid_price is None else bid_price,
                        old_ask=0.0 if ask_price is None else ask_price,
                        new_ask=0.0 if ask_price is None else ask_price
                    )
                    monitor.session.add(snapshot)
    
    # Commit all initial snapshots
    monitor.session.commit()
    print("\n" + "=" * 120)  # Widened to accommodate new columns

def show_recent_captures(minutes=5):
    """Show captures from the last N minutes."""
    session = Session()
    monitor = None
    try:
        # Get captures from last N minutes
        since = datetime.utcnow() - timedelta(minutes=minutes)
        captures = session.query(Snapshot).filter(
            Snapshot.timestamp >= since,
            (Snapshot.old_midpoint != Snapshot.new_midpoint) |  # Show if any price changed
            (Snapshot.old_bid != Snapshot.new_bid) |
            (Snapshot.old_ask != Snapshot.new_ask)
        ).order_by(Snapshot.timestamp.asc()).all()  # Show oldest to newest
        
        if not captures:
            print(f"\nNo changes found in the last {minutes} minutes")
            return
            
        print(f"\nRecent changes (last {minutes} minutes):")
        print("-" * 180)  # Increased width for more columns
        print(f"Time                      | Type | A/D | Spread          | Days |"
              f" Mid Old  | Mid New  |  Mid Δ  |"
              f" Bid Old  | Bid New  |  Bid Δ  |"
              f" Ask Old  | Ask New  |  Ask Δ")
        print("-" * 180)

        # Create and initialize monitor
        monitor = ExcelMonitor()
        monitor.connect_to_excel()  # Connect to Excel
        monitor.update_reference_dates()  # Update reference dates
        
        # Create a list to store all captures with their days
        all_captures = []
        
        # First, build a set of actual spreads from Section 1 (rows 4-29)
        actual_spreads = set()
        if monitor and monitor.excel and monitor.excel.sheet:
            try:
                for row in range(4, 30):  # Rows 4-29 in Section 1
                    date1 = monitor.excel.read_cell(monitor.excel.sheet, f"A{row}")
                    date2 = monitor.excel.read_cell(monitor.excel.sheet, f"B{row}")
                    if date1 and date2:
                        # Format dates consistently
                        if isinstance(date1, datetime):
                            date1 = date1.strftime("%b-%y").upper()
                        if isinstance(date2, datetime):
                            date2 = date2.strftime("%b-%y").upper()
                        actual_spreads.add((str(date1), str(date2)))
            except:
                print("\nWarning: Could not read actual spreads from Excel")
        
        for capture in captures:
            # Calculate changes
            mid_change = capture.new_midpoint - capture.old_midpoint
            bid_change = capture.new_bid - capture.old_bid if not math.isnan(capture.new_bid) and not math.isnan(capture.old_bid) else None
            ask_change = capture.new_ask - capture.old_ask if not math.isnan(capture.new_ask) and not math.isnan(capture.old_ask) else None
            
            # Format change strings
            mid_change_str = f"{capture.old_midpoint:9.2f} -> {capture.new_midpoint:9.2f} ({mid_change:+.2f})"
            
            # Format bid/ask strings - show gray "No Data" if no values or zeros
            if math.isnan(capture.new_bid) or math.isnan(capture.old_bid) or (capture.new_bid == 0 and capture.old_bid == 0):
                bid_change_str = color_text("                 No Data                 ", GRAY)
            else:
                bid_change_str = f"{capture.old_bid:12.2f} -> {capture.new_bid:12.2f} ({bid_change:+.2f})"
                if bid_change > 0:
                    bid_change_str = color_text(bid_change_str, GREEN)
                elif bid_change < 0:
                    bid_change_str = color_text(bid_change_str, RED)
            
            if math.isnan(capture.new_ask) or math.isnan(capture.old_ask) or (capture.new_ask == 0 and capture.old_ask == 0):
                ask_change_str = color_text("                 No Data                 ", GRAY)
            else:
                ask_change_str = f"{capture.old_ask:12.2f} -> {capture.new_ask:12.2f} ({ask_change:+.2f})"
                if ask_change > 0:
                    ask_change_str = color_text(ask_change_str, GREEN)
                elif ask_change < 0:
                    ask_change_str = color_text(ask_change_str, RED)
            
            # Color the mid change
            mid_change_str = f"{capture.old_midpoint:12.2f} -> {capture.new_midpoint:12.2f} ({mid_change:+.2f})"
            if mid_change > 0:
                mid_change_color = GREEN
            elif mid_change < 0:
                mid_change_color = RED
            else:
                mid_change_color = None
            
            # Format mid values
            if mid_change > 0:
                mid_change_str = color_text(mid_change_str, mid_change_color)
            
            # Format bid values
            if math.isnan(capture.old_bid) or math.isnan(capture.new_bid) or (capture.old_bid == 0 and capture.new_bid == 0):
                bid_old_str = color_text("   N/A   ", GRAY)
                bid_new_str = color_text("   N/A   ", GRAY)
                bid_change_str = color_text("  N/A  ", GRAY)
            else:
                bid_old_str = f"{capture.old_bid:10.2f}"
                bid_new_str = f"{capture.new_bid:10.2f}"
                bid_change_str = f"{bid_change:+8.2f}"
                if bid_change > 0:
                    bid_change_str = color_text(bid_change_str, GREEN)
                elif bid_change < 0:
                    bid_change_str = color_text(bid_change_str, RED)
            
            # Format ask values
            if math.isnan(capture.old_ask) or math.isnan(capture.new_ask) or (capture.old_ask == 0 and capture.new_ask == 0):
                ask_old_str = color_text("   N/A   ", GRAY)
                ask_new_str = color_text("   N/A   ", GRAY)
                ask_change_str = color_text("  N/A  ", GRAY)
            else:
                ask_old_str = f"{capture.old_ask:10.2f}"
                ask_new_str = f"{capture.new_ask:10.2f}"
                ask_change_str = f"{ask_change:+8.2f}"
                if ask_change > 0:
                    ask_change_str = color_text(ask_change_str, GREEN)
                elif ask_change < 0:
                    ask_change_str = color_text(ask_change_str, RED)
            
            # Calculate days between for the spread using stored prompts
            days = None
            if capture.prompt1 and capture.prompt2:  # Use stored prompts from database
                if capture.prompt1 == 'C' and monitor.c_date:
                    date1_obj = monitor.c_date
                    date2_obj = get_prompt_date(capture.prompt2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                elif capture.prompt1 == '3M' and monitor.three_m_date:
                    date1_obj = monitor.three_m_date
                    date2_obj = get_prompt_date(capture.prompt2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                elif capture.prompt2 == '3M' and monitor.three_m_date:
                    date1_obj = get_prompt_date(capture.prompt1)
                    date2_obj = monitor.three_m_date
                    if date1_obj:
                        days = abs((date2_obj - date1_obj).days)
                else:
                    days = calculate_days_between(capture.prompt1, capture.prompt2)
            
            days_str = str(days) if days is not None else "N/A"
            
            # Determine if this is an actual spread by checking against our set
            spread_type = "A" if (str(capture.prompt1), str(capture.prompt2)) in actual_spreads else "D"
            
            # Store all capture info for sorting
            all_captures.append((
                capture.timestamp,
                capture.spread_name,
                spread_type,
                days,  # Store actual days for sorting
                days_str,  # Store string version for display
                mid_change_str,
                bid_change_str,
                ask_change_str
            ))
        
        # Sort all captures by days (None values go to the end)
        all_captures.sort(key=lambda x: float('inf') if x[3] is None else x[3])
        
        # Print all captures in a single sorted section
        for (timestamp, spread_name, spread_type, days, days_str, mid_change_str, bid_change_str, ask_change_str) in all_captures:
            # Format mid values
            if mid_change > 0:
                mid_change_color = GREEN
            elif mid_change < 0:
                mid_change_color = RED
            else:
                mid_change_color = None
            
            # Format bid values
            if math.isnan(capture.old_bid) or math.isnan(capture.new_bid) or (capture.old_bid == 0 and capture.new_bid == 0):
                bid_old_str = color_text("   N/A   ", GRAY)
                bid_new_str = color_text("   N/A   ", GRAY)
                bid_change_str = color_text("  N/A  ", GRAY)
            else:
                bid_old_str = f"{capture.old_bid:10.2f}"
                bid_new_str = f"{capture.new_bid:10.2f}"
                bid_change_str = f"{bid_change:+8.2f}"
                if bid_change > 0:
                    bid_change_str = color_text(bid_change_str, GREEN)
                elif bid_change < 0:
                    bid_change_str = color_text(bid_change_str, RED)
            
            # Format ask values
            if math.isnan(capture.old_ask) or math.isnan(capture.new_ask) or (capture.old_ask == 0 and capture.new_ask == 0):
                ask_old_str = color_text("   N/A   ", GRAY)
                ask_new_str = color_text("   N/A   ", GRAY)
                ask_change_str = color_text("  N/A  ", GRAY)
            else:
                ask_old_str = f"{capture.old_ask:10.2f}"
                ask_new_str = f"{capture.new_ask:10.2f}"
                ask_change_str = f"{ask_change:+8.2f}"
                if ask_change > 0:
                    ask_change_str = color_text(ask_change_str, GREEN)
                elif ask_change < 0:
                    ask_change_str = color_text(ask_change_str, RED)
            
            print(f"{timestamp.strftime('%Y-%m-%d %H:%M:%S'):<25} | "
                  f"{spread_name:<15} | "
                  f"{spread_type:^3} | "
                  f"{days_str:>4} | "
                  f"{capture.old_midpoint:10.2f} | "
                  f"{capture.new_midpoint:10.2f} | "
                  f"{color_text(f'{mid_change:+8.2f}', mid_change_color) if mid_change_color else f'{mid_change:+8.2f}'} | "
                  f"{bid_old_str} | {bid_new_str} | {bid_change_str} | "
                  f"{ask_old_str} | {ask_new_str} | {ask_change_str}")
            
        print(f"\nTotal changes: {len(captures)}")
        
    finally:
        session.close()
        if monitor and monitor.excel:
            try:
                if IS_WINDOWS:
                    if hasattr(monitor.excel, 'excel'):
                        monitor.excel.excel.Quit()  # Windows specific cleanup
                else:
                    if hasattr(monitor.excel, 'wb') and hasattr(monitor.excel.wb, 'app'):
                        monitor.excel.wb.app.quit()  # Mac specific cleanup
            except Exception as e:
                print(f"\nWarning: Error during Excel cleanup: {e}")

def capture_with_stability(duration_minutes=None):
    """Run continuous capture with stability checks."""
    start_time = datetime.now()
    monitor = ExcelMonitor()
    
    print(f"\nStarting live price monitoring at {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Stability required: {STABILITY_DURATION} seconds")
    print(f"Minimum price change: {MIN_PRICE_CHANGE}")
    print(f"Poll interval: {POLL_INTERVAL} seconds")
    
    # Take initial snapshot
    print_full_snapshot(monitor, start_time)
    
    print("\nMonitoring for changes:")
    print(f"Time                      | Type | A/D | Spread          | Days |"
          f" Mid Old  | Mid New  |  Mid Δ  |"
          f" Bid Old  | Bid New  |  Bid Δ  |"
          f" Ask Old  | Ask New  |  Ask Δ")
    print("-" * 180)  # Increased width to accommodate wider columns
    
    try:
        while True:
            if duration_minutes and (datetime.now() - start_time).total_seconds() > duration_minutes * 60:
                break
            
            monitor.capture_midpoints()  # Only prints when changes occur
            time.sleep(POLL_INTERVAL)
            
    except KeyboardInterrupt:
        print("\nMonitoring stopped by user")
    finally:
        monitor.session.close()
        
    # Show what was captured
    if duration_minutes:
        show_recent_captures(minutes=duration_minutes)
    else:
        show_recent_captures(minutes=5)

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'show':
        # Just show recent captures without monitoring
        show_recent_captures()
    else:
        # Run continuous monitoring
        capture_with_stability() 