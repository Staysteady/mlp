# Market Maker System

A Python-based market making system that captures and processes market data from Excel, implementing stability logic and historical logging.

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