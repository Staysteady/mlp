# Market Maker System

A Python-based market making system that captures and processes market data from Excel, implementing stability logic and historical logging.

## Quick Start

1. Ensure you have Python 3.10.11 installed:
```bash
# If using pyenv
pyenv install 3.10.11
pyenv local 3.10.11
```

2. Install dependencies:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

3. Configure your Excel file path in `src/market_maker/config/settings.py`:
```python
EXCEL_FILE = "/path/to/your/NEON_ML.xlsm"
```

4. Run the market maker:
```bash
python -m src.market_maker.main
```

## System Components

### 1. Excel Reader
- Reads data from NEON_ML.xlsm (read-only, never writes to Excel)
- Captures:
  - Midpoints from column C
  - Bid/Ask prices from three sections
  - Instrument identifiers for spreads

### 2. Price Monitoring
- Polls Excel every 5 seconds
- Checks price stability (4-second stability required)
- Only logs changes ≥ 0.01
- Operates during trading hours (07:00-16:00)

### 3. Database Logging
- Uses SQLite with WAL enabled
- Stores historical snapshots
- Tracks price changes and spread movements

## CLI Tool Usage

Monitor the system using these commands:

### View Database Statistics
```bash
python -m src.market_maker.cli stats
```

### View Recent Changes
```bash
# Last 5 minutes (default)
python -m src.market_maker.cli recent

# Specify minutes
python -m src.market_maker.cli recent --minutes 10
```

### View Spread History
```bash
# Last 24 hours for specific spread
python -m src.market_maker.cli history JUL24-AUG24

# Specify hours
python -m src.market_maker.cli history JUL24-AUG24 --hours 48
```

### View Largest Price Moves
```bash
# Top 10 moves (default)
python -m src.market_maker.cli moves

# Specify number of moves
python -m src.market_maker.cli moves --top-n 5
```

### View Spread Summary
```bash
# Last 24 hours summary (default)
python -m src.market_maker.cli summary

# Specify hours
python -m src.market_maker.cli summary --hours 12
```

### View Logs
```bash
# View main system logs (default)
python -m src.market_maker.cli logs

# View specific component logs (main, database, or excel)
python -m src.market_maker.cli logs --component database --lines 100
```

## Log Files

The system maintains three log files in the `logs` directory:
- `market_maker.log`: Main system events
- `database.log`: Database operations
- `excel_reader.log`: Excel reading operations

## Configuration

Key settings in `src/market_maker/config/settings.py`:
```python
POLL_INTERVAL = 5          # seconds between Excel polls
STABILITY_DURATION = 4     # seconds price must remain stable
MIN_PRICE_CHANGE = 0.01   # minimum price change to log
TRADING_START_TIME = "07:00"
TRADING_END_TIME = "16:00"
```

## Features

- Excel data capture from specified sheets and ranges
- Price stability monitoring with configurable parameters
- SQLite database with WAL for efficient historical data storage
- Trading hours enforcement (7:00-16:00)
- Graceful shutdown handling
- Configurable settings for easy customization

## Testing

Run the test suite:
```bash
python -m pytest tests/ -v
```

The project uses several testing tools and frameworks:
- `pytest`: Main testing framework
- `freezegun`: Time manipulation for testing time-dependent functionality
  - Used to simulate specific market hours
  - Ensures consistent test results regardless of execution time

Key test areas:
- Excel data reading and parsing
- Database operations and monitoring
- Trading hours validation
- Time utilities and formatting
- Price stability checks

Test files are organized by component:
```
tests/
  unit/
    data/           # Excel reader and database tests
    utils/          # Utility function tests
    models/         # Model tests
```

## Project Structure

```
.
├── NEON_ML.xlsm              # Main Excel data source
├── README.md                 # Project documentation
├── data/                     # Database files
│   ├── market_maker.db
│   ├── market_maker.db-shm
│   └── market_maker.db-wal
├── docs/                     # Additional documentation
├── logs/                     # System log files
│   ├── database.log
│   ├── excel_reader.log
│   └── market_maker.log
├── market_maker/            # Main package directory
│   ├── cli.py              # CLI interface
│   ├── config/             # Configuration files
│   ├── data/               # Data handling modules
│   ├── main.py            # Application entry point
│   ├── models/            # ML/RL model implementations
│   └── utils/             # Helper utilities
├── pyproject.toml         # Project metadata
├── requirements.txt       # Dependencies
├── scripts/              # Utility scripts
│   └── populate_mock_data.py
├── setup.py             # Package setup
└── tests/              # Test suite
    ├── conftest.py    # Test configuration
    ├── test_data/     # Test data files
    └── unit/          # Unit tests by component
```

## Development

- Python 3.10.11 is the recommended version
- The system is designed to be modular and extensible
- Future ML/RL capabilities can be added in the models directory
- Configuration can be modified in settings.py
- Additional features can be added by extending existing classes
- Tests should be added for new functionality using pytest and freezegun where appropriate 