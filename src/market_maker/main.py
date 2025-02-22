"""
Main application module for the market maker system.
Coordinates data capture, processing, and storage.
"""
import time
import schedule
from datetime import datetime
import signal
import sys

from .data.excel_reader import ExcelReader
from .data.models import Session, Snapshot
from .utils.time_utils import is_trading_hours
from .config.settings import (
    POLL_INTERVAL,
    INTERNAL_CHECK_INTERVAL,
    STABILITY_DURATION,
    STARTUP_DELAY,
    MIN_PRICE_CHANGE
)

class MarketMaker:
    """
    Main market maker class that coordinates all system components.
    """
    def __init__(self):
        self.excel_reader = ExcelReader()
        self.session = Session()
        self.running = True
        self.last_snapshot_time = None
        self.stable_count = 0
        
        # Set up signal handlers
        signal.signal(signal.SIGINT, self.handle_shutdown)
        signal.signal(signal.SIGTERM, self.handle_shutdown)

    def handle_shutdown(self, signum, frame):
        """Handle graceful shutdown on signals."""
        print("\nShutting down gracefully...")
        self.running = False
        self.session.close()
        sys.exit(0)

    def process_snapshot(self):
        """
        Process a single snapshot of market data.
        Implements the stability logic and minimum change threshold.
        """
        if not is_trading_hours():
            return

        try:
            # Read current data
            midpoints = self.excel_reader.read_midpoints()
            if midpoints.empty:
                return

            # Check stability
            if self.excel_reader.has_stable_midpoints(midpoints):
                self.stable_count += 1
            else:
                self.stable_count = 0
                return

            # If price has been stable for required duration
            if self.stable_count >= (STABILITY_DURATION / INTERNAL_CHECK_INTERVAL):
                sections_data = self.excel_reader.read_all_sections()
                
                # Process and store changes
                # Implementation will be expanded here
                
                self.stable_count = 0

        except Exception as e:
            print(f"Error processing snapshot: {e}")

    def run(self):
        """
        Main run loop of the market maker system.
        """
        print(f"Starting market maker system in {STARTUP_DELAY} seconds...")
        time.sleep(STARTUP_DELAY)
        
        # Schedule the main processing job
        schedule.every(INTERNAL_CHECK_INTERVAL).seconds.do(self.process_snapshot)
        
        print("Market maker system is running. Press Ctrl+C to stop.")
        
        while self.running:
            schedule.run_pending()
            time.sleep(0.1)

def main():
    """Entry point for the market maker system."""
    market_maker = MarketMaker()
    market_maker.run()

if __name__ == "__main__":
    main() 