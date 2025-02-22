"""
Time-related utility functions for the market maker system.

This module provides utilities for:
- Parsing time strings into time objects
- Checking if current time is within trading hours
- Formatting timestamps for logging
"""
from datetime import datetime, time
from ..config.settings import TRADING_START_TIME, TRADING_END_TIME

def parse_time(time_str: str) -> time:
    """Convert time string to time object.

    Args:
        time_str: Time string in "HH:MM" format

    Returns:
        time: Parsed time object

    Example:
        >>> parse_time("07:00")
        datetime.time(7, 0)
    """
    return datetime.strptime(time_str, "%H:%M").time()

def is_trading_hours() -> bool:
    """Check if current time is within trading hours.

    Trading hours are defined by TRADING_START_TIME and TRADING_END_TIME
    in the settings module (inclusive bounds).

    Returns:
        bool: True if current time is within trading hours
    """
    current_time = datetime.now().time()
    start_time = parse_time(TRADING_START_TIME)
    end_time = parse_time(TRADING_END_TIME)

    return start_time <= current_time <= end_time

def format_timestamp(dt: datetime) -> str:
    """Format datetime for logging.

    Args:
        dt: Datetime object to format

    Returns:
        str: Formatted timestamp string in "YYYY-MM-DD HH:MM:SS" format
    """
    return dt.strftime("%Y-%m-%d %H:%M:%S")
