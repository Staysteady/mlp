"""
Tests for the database monitoring functionality.
"""
import pytest
from datetime import datetime, timedelta
import pandas as pd
from src.market_maker.utils.db_monitor import DatabaseMonitor
from src.market_maker.data.models import Snapshot

@pytest.fixture
def sample_snapshots(db_session):
    """Create sample snapshots for testing."""
    now = datetime.utcnow()
    snapshots = [
        Snapshot(
            timestamp=now - timedelta(minutes=i),
            spread_name=f"JUL24-AUG24",
            prompt1="JUL24",
            prompt2="AUG24",
            old_midpoint=100.0 + i,
            new_midpoint=101.0 + i,
            old_bid=99.5 + i,
            new_bid=100.5 + i,
            old_ask=100.5 + i,
            new_ask=101.5 + i
        ) for i in range(10)
    ]
    
    # Add some data for a different spread
    snapshots.extend([
        Snapshot(
            timestamp=now - timedelta(minutes=i),
            spread_name=f"SEP24-OCT24",
            prompt1="SEP24",
            prompt2="OCT24",
            old_midpoint=102.0 + i,
            new_midpoint=103.0 + i,
            old_bid=101.5 + i,
            new_bid=102.5 + i,
            old_ask=102.5 + i,
            new_ask=103.5 + i
        ) for i in range(5)
    ])
    
    for snapshot in snapshots:
        db_session.add(snapshot)
    db_session.commit()
    
    return snapshots

class TestDatabaseMonitor:
    def test_get_recent_snapshots(self, db_session, sample_snapshots):
        """Test retrieving recent snapshots."""
        monitor = DatabaseMonitor(db_session)
        
        # Get snapshots from last 5 minutes
        recent = monitor.get_recent_snapshots(minutes=5)
        assert len(recent) == 10  # Should get both spreads' snapshots within 5 mins
        
        # Get snapshots from last 2 minutes
        very_recent = monitor.get_recent_snapshots(minutes=2)
        assert len(very_recent) == 4  # Should get fewer snapshots
    
    def test_get_spread_history(self, db_session, sample_snapshots):
        """Test retrieving spread history."""
        monitor = DatabaseMonitor(db_session)
        
        # Get history for first spread
        history = monitor.get_spread_history("JUL24-AUG24", hours=24)
        assert not history.empty
        assert len(history) == 10
        assert all(history['new_mid'] > history['old_mid'])
        
        # Test non-existent spread
        empty_history = monitor.get_spread_history("NON-EXISTENT", hours=24)
        assert empty_history.empty
    
    def test_get_database_stats(self, db_session, sample_snapshots):
        """Test retrieving database statistics."""
        monitor = DatabaseMonitor(db_session)
        stats = monitor.get_database_stats()
        
        assert stats['total_snapshots'] == 15  # Total from sample data
        assert stats['unique_spreads'] == 2    # Two different spreads
        assert isinstance(stats['oldest_record'], datetime)
        assert isinstance(stats['newest_record'], datetime)
    
    def test_get_largest_moves(self, db_session, sample_snapshots):
        """Test retrieving largest price moves."""
        monitor = DatabaseMonitor(db_session)
        moves = monitor.get_largest_moves(top_n=5)
        
        assert len(moves) == 5
        assert 'change' in moves.columns
        # Verify moves are sorted by size
        changes = moves['change'].tolist()
        assert changes == sorted(changes, reverse=True)
    
    def test_get_spread_summary(self, db_session, sample_snapshots):
        """Test retrieving spread summary."""
        monitor = DatabaseMonitor(db_session)
        summary = monitor.get_spread_summary(hours=24)
        
        assert len(summary) == 2  # Two different spreads
        assert 'updates' in summary.columns
        assert 'avg_change' in summary.columns
        assert 'min_price' in summary.columns
        assert 'max_price' in summary.columns
        
        # Verify update counts
        spread_counts = summary.set_index(summary.columns[0])['updates']
        assert spread_counts['JUL24-AUG24'] == 10
        assert spread_counts['SEP24-OCT24'] == 5
    
    def test_context_manager(self, db_session):
        """Test the context manager functionality."""
        with DatabaseMonitor(db_session) as monitor:
            stats = monitor.get_database_stats()
            assert isinstance(stats, dict)
        # Session should be closed after context manager exits
        with pytest.raises(Exception):
            # This should raise an error since the session is closed
            db_session.execute("SELECT 1") 