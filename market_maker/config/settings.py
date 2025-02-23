"""
Configuration settings for the market maker system.
All configurable parameters are centralized here for easy modification.
"""
from pathlib import Path
import os

# Get project root directory
PROJECT_ROOT = Path(__file__).parent.parent.parent

# Excel file settings
EXCEL_FILE = str(PROJECT_ROOT / "NEON_ML.xlsm")
SHEET_NAME = "AH NEON"
PREFIX_CELL = "B2"  # Cell containing the spread prefix (e.g., "AHD")

# Cell ranges for data capture
MIDPOINT_RANGE = "C4:C29"

# Section 1 configuration (columns A-G)
SECTION1_RANGES = {
    "prompt1_col": "A",  # First prompt column
    "prompt2_col": "B",  # Second prompt column
    "bid": "F",
    "ask": "G"
}

# Section 2 configuration (columns Z-AF)
SECTION2_RANGES = {
    "prompt1_col": "Z",  # First prompt column
    "prompt2_col": "AA", # Second prompt column
    "bid": "AE",
    "ask": "AF"
}

# Section 3 configuration (columns AW-BC)
SECTION3_RANGES = {
    "prompt1_col": "AW", # First prompt column
    "prompt2_col": "AX", # Second prompt column
    "bid": "BB",
    "ask": "BC"
}

# Timing settings
POLL_INTERVAL = 5  # seconds between Excel polls
INTERNAL_CHECK_INTERVAL = 1  # seconds between internal price checks
STABILITY_DURATION = 4  # seconds price must remain stable
STARTUP_DELAY = 10  # seconds to wait before starting

# Trading hours
TRADING_START_TIME = "07:00"
TRADING_END_TIME = "16:00"

# Price thresholds
MIN_PRICE_CHANGE = 0.01  # minimum price change to log

# Database settings
DB_PATH = PROJECT_ROOT / "data" / "market_maker.db"
ENABLE_WAL = True  # Write-Ahead Logging for better performance

# Create data directory if it doesn't exist
os.makedirs(DB_PATH.parent, exist_ok=True) 