"""
Script to view recent captures from the database.
"""
from market_maker.data.models import Session, Snapshot
from datetime import datetime, timedelta
import argparse

def show_recent_captures(minutes=5):
    """Show captures from the last N minutes."""
    session = Session()
    try:
        # Get captures from last N minutes
        since = datetime.utcnow() - timedelta(minutes=minutes)
        captures = session.query(Snapshot).filter(
            Snapshot.timestamp >= since
        ).order_by(Snapshot.timestamp.desc()).all()
        
        if not captures:
            print(f"\nNo captures found in the last {minutes} minutes")
            return
            
        print(f"\nRecent captures (last {minutes} minutes):")
        print("-" * 100)
        print(f"{'Timestamp':<20} | {'Spread':<15} | {'Old Mid':>9} | {'New Mid':>9} | {'Change':>9}")
        print("-" * 100)
        
        for capture in captures:
            change = capture.new_midpoint - capture.old_midpoint
            print(f"{capture.timestamp.strftime('%H:%M:%S'):<20} | "
                  f"{capture.spread_name:<15} | "
                  f"{capture.old_midpoint:9.2f} | "
                  f"{capture.new_midpoint:9.2f} | "
                  f"{change:+9.2f}")
            
        print(f"\nTotal captures: {len(captures)}")
        
    finally:
        session.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='View recent captures from the database')
    parser.add_argument('--minutes', type=int, default=5, help='Show captures from last N minutes')
    args = parser.parse_args()
    
    show_recent_captures(minutes=args.minutes) 