"""
Database monitoring utilities for the market maker system.
Provides functions to query and monitor database state.
"""
from typing import List, Optional
from datetime import datetime, timedelta
import pandas as pd
from sqlalchemy import select, func
from ..data.models import Snapshot, Session
from .logging_config import db_logger

class DatabaseMonitor:
    """
    Utility class for monitoring and querying the database.
    Provides methods to view recent changes and database statistics.
    """
    def __init__(self, session: Optional[Session] = None):
        self.session = session or Session()
        self.logger = db_logger
    
    def get_recent_snapshots(self, minutes: int = 5) -> List[Snapshot]:
        """Get snapshots from the last N minutes."""
        cutoff_time = datetime.utcnow() - timedelta(minutes=minutes)
        query = select(Snapshot).where(Snapshot.timestamp >= cutoff_time)
        
        snapshots = self.session.execute(query).scalars().all()
        self.logger.info(f"Retrieved {len(snapshots)} snapshots from last {minutes} minutes")
        return snapshots
    
    def get_spread_history(self, spread_name: str, hours: int = 24) -> pd.DataFrame:
        """Get price history for a specific spread."""
        cutoff_time = datetime.utcnow() - timedelta(hours=hours)
        query = select(Snapshot).where(
            Snapshot.spread_name == spread_name,
            Snapshot.timestamp >= cutoff_time
        )
        
        snapshots = self.session.execute(query).scalars().all()
        
        # Convert to DataFrame for easier analysis
        if not snapshots:
            self.logger.warning(f"No history found for spread {spread_name}")
            return pd.DataFrame()
        
        data = {
            'timestamp': [s.timestamp for s in snapshots],
            'old_mid': [s.old_midpoint for s in snapshots],
            'new_mid': [s.new_midpoint for s in snapshots],
            'old_bid': [s.old_bid for s in snapshots],
            'new_bid': [s.new_bid for s in snapshots],
            'old_ask': [s.old_ask for s in snapshots],
            'new_ask': [s.new_ask for s in snapshots]
        }
        
        df = pd.DataFrame(data)
        self.logger.info(f"Retrieved {len(df)} historical records for {spread_name}")
        return df
    
    def get_database_stats(self) -> dict:
        """Get general database statistics."""
        stats = {
            'total_snapshots': self.session.scalar(select(func.count(Snapshot.id))),
            'unique_spreads': self.session.scalar(
                select(func.count(func.distinct(Snapshot.spread_name)))
            ),
            'oldest_record': self.session.scalar(
                select(func.min(Snapshot.timestamp))
            ),
            'newest_record': self.session.scalar(
                select(func.max(Snapshot.timestamp))
            )
        }
        
        self.logger.info(f"Database stats: {stats}")
        return stats
    
    def get_largest_moves(self, top_n: int = 10) -> pd.DataFrame:
        """Get the largest price moves in the database."""
        query = select(Snapshot).order_by(
            (Snapshot.new_midpoint - Snapshot.old_midpoint).desc()
        ).limit(top_n)
        
        moves = self.session.execute(query).scalars().all()
        
        data = {
            'timestamp': [m.timestamp for m in moves],
            'spread': [m.spread_name for m in moves],
            'old_mid': [m.old_midpoint for m in moves],
            'new_mid': [m.new_midpoint for m in moves],
            'change': [m.new_midpoint - m.old_midpoint for m in moves]
        }
        
        df = pd.DataFrame(data)
        self.logger.info(f"Retrieved top {top_n} largest price moves")
        return df

    def get_spread_summary(self, hours: int = 24) -> pd.DataFrame:
        """Get a summary of all spreads' activity."""
        cutoff_time = datetime.utcnow() - timedelta(hours=hours)
        query = select(
            Snapshot.spread_name,
            func.count(Snapshot.id).label('updates'),
            func.avg(Snapshot.new_midpoint - Snapshot.old_midpoint).label('avg_change'),
            func.min(Snapshot.new_midpoint).label('min_price'),
            func.max(Snapshot.new_midpoint).label('max_price')
        ).where(
            Snapshot.timestamp >= cutoff_time
        ).group_by(
            Snapshot.spread_name
        )
        
        result = self.session.execute(query).all()
        
        if not result:
            self.logger.warning(f"No spread activity in the last {hours} hours")
            return pd.DataFrame()
        
        df = pd.DataFrame(result)
        self.logger.info(f"Generated summary for {len(df)} spreads")
        return df
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            # If no exception occurred, commit any pending changes
            self.session.commit()
        else:
            # If an exception occurred, rollback
            self.session.rollback()
        
        self.session.close()
        return False  # Re-raise any exceptions 