# Market Maker System

A Python-based market making system that captures and processes market data from Excel, implementing stability logic and historical logging.

## Quick Start

1. Install dependencies:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

2. Configure your Excel file path in `src/market_maker/config/settings.py`:
```python
EXCEL_FILE = "/path/to/your/NEON_ML.xlsm"
```

3. Run the market maker:
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

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure settings in `src/market_maker/config/settings.py`:
- Excel file path and sheet name
- Cell ranges for data capture
- Timing parameters
- Trading hours
- Database settings

## Running the System

To start the market maker:

```bash
python -m src.market_maker.main
```

The system will:
1. Wait for the configured startup delay
2. Begin monitoring Excel data during trading hours
3. Log stable price snapshots to the SQLite database
4. Handle graceful shutdown on Ctrl+C

## Project Structure

```
src/
  market_maker/
    data/           # Excel reader and database operations
    models/         # Future ML/RL models
    utils/          # Helper functions
    config/         # Configuration management
tests/             # Unit tests
docs/              # Documentation
```

## Development

- The system is designed to be modular and extensible
- Future ML/RL capabilities can be added in the models directory
- Configuration can be modified in settings.py
- Additional features can be added by extending existing classes 