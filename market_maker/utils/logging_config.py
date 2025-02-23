"""
Logging configuration for the market maker system.
Implements both file and console logging with different levels.
"""
import logging
import sys
from pathlib import Path
from logging.handlers import RotatingFileHandler
import os

# Create logs directory if it doesn't exist
LOGS_DIR = Path("logs")
LOGS_DIR.mkdir(exist_ok=True)

# Create separate log files for different components
MAIN_LOG = LOGS_DIR / "market_maker.log"
DB_LOG = LOGS_DIR / "database.log"
EXCEL_LOG = LOGS_DIR / "excel_reader.log"

def setup_logger(name: str, log_file: Path, level=logging.INFO):
    """Set up a logger with both file and console handlers."""
    level_env = os.getenv("MARKET_MAKER_LOG_LEVEL")
    if level_env:
        try:
            level = getattr(logging, level_env.upper(), level)
        except Exception:
            pass
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Prevent duplicate handlers
    if logger.handlers:
        return logger

    # Create formatters
    file_formatter = logging.Formatter(
        '%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_formatter = logging.Formatter(
        '%(asctime)s | %(levelname)-8s | %(message)s',
        datefmt='%H:%M:%S'
    )

    # File handler (rotating log file)
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
    file_handler.setFormatter(file_formatter)
    file_handler.setLevel(logging.DEBUG)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(console_formatter)
    console_handler.setLevel(logging.INFO)

    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

# Create loggers for different components
main_logger = setup_logger('market_maker', MAIN_LOG)
db_logger = setup_logger('database', DB_LOG)
excel_logger = setup_logger('excel_reader', EXCEL_LOG)
