"""
Main application module for the market maker system.
Coordinates data capture, processing, and storage.
"""
# Standard library imports
import signal
import sys
import time

# Third-party imports
import schedule

# Local imports
from .data.excel_reader import ExcelReader
from .data.models import Session
from .utils.time_utils import is_trading_hours
from .utils.logging_config import main_logger
from .utils.db_monitor import DatabaseMonitor
from .config.settings import (
    INTERNAL_CHECK_INTERVAL,
    STABILITY_DURATION,
    STARTUP_DELAY
)

class MarketMaker:
    """
    Main market maker class that coordinates all system components.
    """
    def __init__(self):
        self.excel_reader = ExcelReader()
        self.session = Session()
        self.db_monitor = DatabaseMonitor(self.session)
        self.running = True
        self.last_snapshot_time = None
        self.stable_count = 0
        self.logger = main_logger

        # Set up signal handlers
        signal.signal(signal.SIGINT, self.handle_shutdown)
        signal.signal(signal.SIGTERM, self.handle_shutdown)

        self.logger.info("Market maker system initialized")

    def handle_shutdown(self, signum, frame):  # pylint: disable=unused-argument
        """Handle graceful shutdown on signals."""
        self.logger.info("Shutting down gracefully...")
        self.running = False
        self.session.close()
        sys.exit(0)

    def process_snapshot(self):
        """
        Process a single snapshot of market data.
        Implements the stability logic and minimum change threshold.
        """
        if not is_trading_hours():
            self.logger.debug("Outside trading hours, skipping snapshot")
            return

        try:
            # Read current data
            midpoints = self.excel_reader.read_midpoints()
            if midpoints.empty:
                self.logger.warning("Failed to read midpoints")
                return

            # Check stability
            if self.excel_reader.has_stable_midpoints(midpoints):
                self.stable_count += 1
                self.logger.debug("Prices stable for %(count)s checks", 
                                {'count': self.stable_count})
            else:
                self.stable_count = 0
                self.logger.debug("Price change detected, resetting stability counter")
                return

            # If price has been stable for required duration
            if self.stable_count >= (STABILITY_DURATION / INTERNAL_CHECK_INTERVAL):
                self.logger.info("Price stability threshold reached, processing snapshot")

                # Process and store changes
                # Implementation will be expanded here

                # Log database stats periodically
                stats = self.db_monitor.get_database_stats()
                self.logger.info("Database stats after snapshot: %(stats)s", 
                               {'stats': stats})

                self.stable_count = 0

        except (ValueError, KeyError) as e:
            self.logger.error("Error processing snapshot: %(error)s", 
                            {'error': str(e)}, exc_info=True)

    def run(self):
        """
        Main run loop of the market maker system.
        """
        self.logger.info("Starting market maker system in %(delay)s seconds...",
                        {'delay': STARTUP_DELAY})
        time.sleep(STARTUP_DELAY)

        # Schedule the main processing job
        schedule.every(INTERNAL_CHECK_INTERVAL).seconds.do(self.process_snapshot)

        self.logger.info("Market maker system is running. Press Ctrl+C to stop.")

        while self.running:
            schedule.run_pending()
            time.sleep(0.1)

def main():
    """Entry point for the market maker system."""
    market_maker = MarketMaker()
    market_maker.run()

if __name__ == "__main__":
    main() 