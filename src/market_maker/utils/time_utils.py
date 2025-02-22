"""
Time-related utility functions for the market maker system.
"""
from datetime import datetime, time
from ..config.settings import TRADING_START_TIME, TRADING_END_TIME

def parse_time(time_str: str) -> time:
    """Convert time string to time object."""
    return datetime.strptime(time_str, "%H:%M").time()

def is_trading_hours() -> bool:
    """Check if current time is within trading hours."""
    current_time = datetime.now().time()
    start_time = parse_time(TRADING_START_TIME)
    end_time = parse_time(TRADING_END_TIME)
    
    return start_time <= current_time <= end_time

def format_timestamp(dt: datetime) -> str:
    """Format datetime for logging."""
    return dt.strftime("%Y-%m-%d %H:%M:%S") 