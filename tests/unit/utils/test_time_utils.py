"""
Tests for time-related utility functions.
"""
import pytest
from datetime import datetime, time
from freezegun import freeze_time
from src.market_maker.utils.time_utils import parse_time, is_trading_hours, format_timestamp

class TestTimeUtils:
    def test_parse_time(self):
        """Test parsing time strings."""
        parsed = parse_time("07:00")
        assert isinstance(parsed, time)
        assert parsed.hour == 7
        assert parsed.minute == 0
        
        parsed = parse_time("16:00")
        assert parsed.hour == 16
        assert parsed.minute == 0
    
    @freeze_time("2024-03-20 08:30:00")
    def test_is_trading_hours_during_trading(self):
        """Test trading hours check during trading time."""
        assert is_trading_hours() is True
    
    @freeze_time("2024-03-20 06:59:59")
    def test_is_trading_hours_before_trading(self):
        """Test trading hours check before trading starts."""
        assert is_trading_hours() is False
    
    @freeze_time("2024-03-20 16:00:01")
    def test_is_trading_hours_after_trading(self):
        """Test trading hours check after trading ends."""
        assert is_trading_hours() is False
    
    def test_format_timestamp(self):
        """Test timestamp formatting."""
        dt = datetime(2024, 3, 20, 14, 30, 45)
        formatted = format_timestamp(dt)
        assert formatted == "2024-03-20 14:30:45" 