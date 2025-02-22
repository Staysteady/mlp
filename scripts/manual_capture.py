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

# Initialize database
engine = init_db()

# ANSI color codes
GREEN = '\033[32m'
RED = '\033[31m'
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
    if any(x is None for x in [date1, date2, value]):
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
        if math.isnan(float_value):  # Only filter NaN, allow zero values
            return False
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
    def __init__(self, spread, value, timestamp, is_primary=False):
        self.spread = spread
        self.value = self._safe_float_conversion(value)
        self.first_seen = timestamp
        self.last_seen = timestamp
        self.unchanged_since = timestamp
        self.is_stable = False  # Start as unstable
        self.is_recorded = False  # Track if we've recorded this value
        self.last_recorded_value = self._safe_float_conversion(value)  # Initialize with current value
        self.is_primary = is_primary  # Whether this is a primary spread (Section 1, rows 4-29)
        self.dependency = None  # Track which primary spread caused this change

    def _safe_float_conversion(self, value):
        """Safely convert a value to float, handling various types."""
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            try:
                return float(value)
            except ValueError:
                return 0.0
        return 0.0  # Default for any other type (datetime, etc)
    
    def update(self, value, timestamp, dependency=None):
        """Update price point with new value."""
        try:
            current_value = self._safe_float_conversion(value)
            if abs(self.value - current_value) >= MIN_PRICE_CHANGE:
                # Price changed significantly, reset stability tracking
                old_value = self.last_recorded_value  # Use last recorded value instead of current
                self.value = current_value
                self.unchanged_since = timestamp
                self.is_stable = False
                self.is_recorded = False  # Reset recorded flag on significant change
                if dependency:
                    self.dependency = dependency  # Update dependency if provided
                return False, old_value
            else:
                # Price hasn't changed significantly
                self.last_seen = timestamp
                # Check if price has been stable for required duration
                stability_duration = (timestamp - self.unchanged_since).total_seconds()
                was_stable = self.is_stable
                self.is_stable = stability_duration >= STABILITY_DURATION
                # Return True only when we first become stable and haven't recorded yet
                return (self.is_stable and not was_stable and not self.is_recorded), self.last_recorded_value
        except (ValueError, TypeError) as e:
            print(f"Error updating price point: {e}")
            return False, self.last_recorded_value

    def mark_recorded(self):
        """Mark this price point as having been recorded."""
        self.is_recorded = True
        self.last_recorded_value = self.value  # Update last recorded value

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
            self.wb = xw.books['NEON_ML.xlsm']
            self.sheet = self.wb.sheets["AH NEON"]
            self.sod_sheet = self.wb.sheets["SOD"]
        except Exception as e:
            print(f"Error connecting to Excel on Mac: {e}")
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
        self.connect_to_excel()
        self.update_reference_dates()
    
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
            changes = []  # Track changes for printing
            capture_time = datetime.utcnow()
            
            # Section 1 (A-C)
            for row in range(2, 100):  # Adjust range as needed
                date1 = self.excel.read_cell(self.excel.sheet, f"A{row}")
                date2 = self.excel.read_cell(self.excel.sheet, f"B{row}")
                mid = self.excel.read_cell(self.excel.sheet, f"C{row}")
                
                self._process_spread_data(date1, date2, mid, seen_spreads, changes, capture_time, current_values)

            # Section 2 (Z-AB)
            for row in range(2, 100):
                date1 = self.excel.read_cell(self.excel.sheet, f"Z{row}")
                date2 = self.excel.read_cell(self.excel.sheet, f"AA{row}")
                mid = self.excel.read_cell(self.excel.sheet, f"AB{row}")
                
                self._process_spread_data(date1, date2, mid, seen_spreads, changes, capture_time, current_values)

            # Section 3 (AW-AY)
            for row in range(2, 100):
                date1 = self.excel.read_cell(self.excel.sheet, f"AW{row}")
                date2 = self.excel.read_cell(self.excel.sheet, f"AX{row}")
                mid = self.excel.read_cell(self.excel.sheet, f"AY{row}")
                
                self._process_spread_data(date1, date2, mid, seen_spreads, changes, capture_time, current_values)
            
            # Commit all changes to database
            if changes:
                self.session.commit()
                
                # Print changes
                print("\nTime                      | Type | Spread          | Days | Value/Change")
                print("-" * 80)
                for change_data in changes:
                    if len(change_data) == 7:  # Has dependency info and days
                        time, type_, spread, new_value, old_value, dependency, days = change_data
                    else:
                        time, type_, spread, new_value, old_value, dependency = change_data
                        days = None
                    
                    days_str = str(days) if days is not None else "N/A"
                    
                    change = new_value - old_value
                    change_str = f"{change:+.2f}"
                    if change > 0:
                        change_str = color_text(change_str, GREEN)
                    elif change < 0:
                        change_str = color_text(change_str, RED)
                    
                    # Add dependency info if available
                    if dependency:
                        print(f"{time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} {type_} {spread:<15s} | {days_str:>4} | {old_value:9.2f} -> {new_value:9.2f} ({change_str}) [from {dependency}]")
                    else:
                        print(f"{time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} {type_} {spread:<15s} | {days_str:>4} | {old_value:9.2f} -> {new_value:9.2f} ({change_str})")
            
            return current_values
        except Exception as e:
            print(f"Error capturing midpoints: {e}")
            self.session.rollback()  # Rollback on error
            return {}

    def _process_spread_data(self, date1, date2, mid, seen_spreads, changes, capture_time, current_values):
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
                spread = f"{date1}-{date2}"
                
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
                is_primary = False
                primary_dependency = None
                
                # Check if any primary spreads have changed recently
                for key, point in self.price_tracker.items():
                    if point.is_primary and not point.is_recorded and point.unchanged_since >= capture_time - timedelta(seconds=POLL_INTERVAL):
                        primary_dependency = key
                        break
                
                # Check if this is a primary spread (Section 1, rows 4-29)
                if spread_key.startswith('A') and 4 <= int(spread_key.split('-')[1]) <= 29:
                    is_primary = True
                
                # Check if this is a new spread or value has changed
                if spread not in self.price_tracker:
                    # New spread - start tracking but don't show in changes
                    self.price_tracker[spread] = PricePoint(spread, current_value, capture_time, is_primary=is_primary)
                    if current_value != 0:  # Only record non-zero values
                        snapshot = Snapshot(
                            timestamp=capture_time,
                            spread_name=spread,
                            prompt1=date1,
                            prompt2=date2,
                            old_midpoint=current_value,
                            new_midpoint=current_value,
                            old_bid=0.0,
                            new_bid=0.0,
                            old_ask=0.0,
                            new_ask=0.0
                        )
                        self.session.add(snapshot)
                        self.price_tracker[spread].mark_recorded()
                else:
                    price_point = self.price_tracker[spread]
                    should_record, old_value = price_point.update(current_value, capture_time, dependency=primary_dependency)
                    
                    # Check if value has changed significantly
                    if abs(current_value - price_point.last_recorded_value) >= MIN_PRICE_CHANGE:
                        # Value has changed significantly
                        if current_value != 0:  # Only show non-zero values
                            # Record the change after stability period
                            if should_record:
                                # Add dependency info to changes list if this is a derived spread
                                if not price_point.is_primary and price_point.dependency:
                                    changes.append((capture_time, "CHG", spread, current_value, price_point.last_recorded_value, price_point.dependency, days))
                                else:
                                    changes.append((capture_time, "CHG", spread, current_value, price_point.last_recorded_value, None, days))
                                
                                snapshot = Snapshot(
                                    timestamp=capture_time,
                                    spread_name=spread,
                                    prompt1=date1,
                                    prompt2=date2,
                                    old_midpoint=price_point.last_recorded_value,
                                    new_midpoint=current_value,
                                    old_bid=0.0,
                                    new_bid=0.0,
                                    old_ask=0.0,
                                    new_ask=0.0
                                )
                                self.session.add(snapshot)
                                price_point.mark_recorded()
                
                current_values[spread] = current_value
                seen_spreads.add(spread_key)
            except (ValueError, TypeError) as e:
                print(f"Warning: Invalid value for spread {date1}-{date2}: {mid}")

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
        ("Section 1", "A", "B", "C"),
        ("Section 2", "Z", "AA", "AB"),
        ("Section 3", "AW", "AX", "AY")
    ]
    
    # Initialize price points for monitoring without waiting for stability
    for section_name, col1, col2, col3 in sections:
        print(f"\n{section_name}")
        print("-" * 120)  # Widened to accommodate new columns
        print(f"{'Spread':<15} | {'Midpoint':>10} | {'Days Between':>12} | {'Dates':>30}")
        print("-" * 120)  # Widened to accommodate new columns
        
        for row in range(4, 100):  # Start from row 4 where data begins
            date1 = monitor.excel.read_cell(monitor.excel.sheet, f"{col1}{row}")
            date2 = monitor.excel.read_cell(monitor.excel.sheet, f"{col2}{row}")
            mid = monitor.excel.read_cell(monitor.excel.sheet, f"{col3}{row}")
            
            if is_valid_spread(date1, date2, mid):
                if isinstance(date1, datetime):
                    date1 = date1.strftime("%b-%y").upper()
                if isinstance(date2, datetime):
                    date2 = date2.strftime("%b-%y").upper()
                spread = f"{date1}-{date2}"
                
                # Calculate days between prompts
                days = None
                dates_str = ""
                
                # Handle special cases with C and 3M
                if date1 == 'C' and monitor.c_date:
                    date1_obj = monitor.c_date
                    date2_obj = get_prompt_date(date2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                        dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                elif date1 == '3M' and monitor.three_m_date:
                    date1_obj = monitor.three_m_date
                    date2_obj = get_prompt_date(date2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                        dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                elif date2 == '3M' and monitor.three_m_date:
                    date1_obj = get_prompt_date(date1)
                    date2_obj = monitor.three_m_date
                    if date1_obj:
                        days = abs((date2_obj - date1_obj).days)
                        dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                else:
                    days = calculate_days_between(date1, date2)
                    if days is not None:
                        date1_obj = get_prompt_date(date1)
                        date2_obj = get_prompt_date(date2)
                        if date1_obj and date2_obj:
                            dates_str = f"{date1_obj.strftime('%Y-%m-%d')} → {date2_obj.strftime('%Y-%m-%d')}"
                
                days_str = str(days) if days is not None else "N/A"
                
                print(f"{spread:<15} | {float(mid):10.2f} | {days_str:>12} | {dates_str:>30}")
                
                # Initialize price point for monitoring but mark as recorded
                if spread not in monitor.price_tracker:
                    price_point = PricePoint(spread, mid, capture_time)
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
                        old_bid=0.0,
                        new_bid=0.0,
                        old_ask=0.0,
                        new_ask=0.0
                    )
                    monitor.session.add(snapshot)
    
    # Commit all initial snapshots
    monitor.session.commit()
    print("\n" + "=" * 120)  # Widened to accommodate new columns

def show_recent_captures(minutes=5):
    """Show captures from the last N minutes."""
    session = Session()
    try:
        # Get captures from last N minutes
        since = datetime.utcnow() - timedelta(minutes=minutes)
        captures = session.query(Snapshot).filter(
            Snapshot.timestamp >= since,
            Snapshot.old_midpoint != Snapshot.new_midpoint  # Only show actual changes
        ).order_by(Snapshot.timestamp.asc()).all()  # Show oldest to newest
        
        if not captures:
            print(f"\nNo changes found in the last {minutes} minutes")
            return
            
        print(f"\nRecent changes (last {minutes} minutes):")
        print("-" * 120)
        print(f"{'Timestamp':<25} | {'Spread':<15} | {'Days':>4} | {'Old Mid':>9} | {'New Mid':>9} | {'Change':>9}")
        print("-" * 120)
        
        monitor = ExcelMonitor()  # Create monitor to access reference dates
        
        for capture in captures:
            change = capture.new_midpoint - capture.old_midpoint
            change_str = f"{change:+9.2f}"
            if change > 0:
                change_str = color_text(change_str, GREEN)
            elif change < 0:
                change_str = color_text(change_str, RED)
                
            # Calculate days between for the spread - split on first hyphen only
            parts = capture.spread_name.split('-', 1)  # Split on first hyphen only
            if len(parts) == 2:
                date1, date2 = parts
                days = None
                
                # Handle special cases with C and 3M
                if date1 == 'C' and monitor.c_date:
                    date1_obj = monitor.c_date
                    date2_obj = get_prompt_date(date2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                elif date1 == '3M' and monitor.three_m_date:
                    date1_obj = monitor.three_m_date
                    date2_obj = get_prompt_date(date2)
                    if date2_obj:
                        days = abs((date2_obj - date1_obj).days)
                elif date2.endswith('3M') and monitor.three_m_date:  # Handle cases where 3M is at the end
                    date1_obj = get_prompt_date(date1)
                    date2_obj = monitor.three_m_date
                    if date1_obj:
                        days = abs((date2_obj - date1_obj).days)
                else:
                    days = calculate_days_between(date1, date2)
                
                days_str = str(days) if days is not None else "N/A"
            else:
                days_str = "N/A"
            
            print(f"{capture.timestamp.strftime('%Y-%m-%d %H:%M:%S'):<25} | "
                  f"{capture.spread_name:<15} | "
                  f"{days_str:>4} | "
                  f"{capture.old_midpoint:9.2f} | "
                  f"{capture.new_midpoint:9.2f} | "
                  f"{change_str}")
            
        print(f"\nTotal changes: {len(captures)}")
        
    finally:
        session.close()

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
    print(f"{'Time':<25} | {'Type':<4s} | {'Spread':<15s} | {'Value/Change':<30s}")
    print("-" * 80)
    
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