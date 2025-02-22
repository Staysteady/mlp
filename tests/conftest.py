"""
Pytest configuration and fixtures.

Note on Test Data vs Production:
------------------------------
The market maker system only reads from Excel files - it never writes to them.
All tests use mocked data to simulate Excel reading operations, ensuring we
maintain the read-only nature of the system even in our tests.
"""
from pathlib import Path

import pandas as pd
import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from market_maker.data.models import Base, Snapshot

# Test database path
TEST_DB_PATH = Path(__file__).parent / "test_data/test.db"

@pytest.fixture(scope="session")
def test_db_engine():
    """Create a test database engine."""
    # Create test data directory if it doesn't exist
    TEST_DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Create test database
    engine = create_engine(f"sqlite:///{TEST_DB_PATH}")
    Base.metadata.create_all(engine)

    yield engine

    # Cleanup
    if TEST_DB_PATH.exists():
        TEST_DB_PATH.unlink()

@pytest.fixture
def db_session(test_db_engine):
    """Create a new database session for a test."""
    session_factory = sessionmaker(bind=test_db_engine)
    session = session_factory()

    # Clean up any existing data before test
    session.query(Snapshot).delete()
    session.commit()

    yield session

    # Clean up all data and close session after test
    session.query(Snapshot).delete()
    session.commit()
    session.close()

@pytest.fixture
def mock_midpoints_data():
    """Return mock data for midpoints."""
    return pd.Series([100.25, 101.25, 102.25], name="C")

@pytest.fixture
def mock_section1_data():
    """Return mock data for section 1."""
    prompts = pd.DataFrame({
        'prompt1': ["JUL24", "AUG24", "SEP24"],
        'prompt2': ["AUG24", "SEP24", "OCT24"]
    })
    prices = pd.DataFrame({
        'bid': [100.0, 101.0, 102.0],
        'ask': [100.5, 101.5, 102.5]
    })
    return prompts, prices

@pytest.fixture
def mock_section2_data():
    """Return mock data for section 2."""
    prompts = pd.DataFrame({
        'prompt1': ["OCT24", "NOV24", "DEC24"],
        'prompt2': ["NOV24", "DEC24", "JAN25"]
    })
    prices = pd.DataFrame({
        'bid': [103.0, 104.0, 105.0],
        'ask': [103.5, 104.5, 105.5]
    })
    return prompts, prices

@pytest.fixture
def mock_section3_data():
    """Return mock data for section 3."""
    prompts = pd.DataFrame({
        'prompt1': ["JAN25", "FEB25", "MAR25"],
        'prompt2': ["FEB25", "MAR25", "APR25"]
    })
    prices = pd.DataFrame({
        'bid': [106.0, 107.0, 108.0],
        'ask': [106.5, 107.5, 108.5]
    })
    return prompts, prices
