"""
Script to populate the database with mock data for testing.
"""
from datetime import datetime, timedelta
from src.market_maker.data.models import Session, Snapshot

def populate_mock_data():
    """Populate database with mock data."""
    session = Session()
    now = datetime.utcnow()
    
    # Create mock data for different spreads
    spreads = [
        {
            'name': 'JUL24-AUG24',
            'prompt1': 'JUL24',
            'prompt2': 'AUG24',
            'base_mid': 100.0,
            'base_bid': 99.5,
            'base_ask': 100.5
        },
        {
            'name': 'SEP24-OCT24',
            'prompt1': 'SEP24',
            'prompt2': 'OCT24',
            'base_mid': 102.0,
            'base_bid': 101.5,
            'base_ask': 102.5
        },
        {
            'name': 'NOV24-DEC24',
            'prompt1': 'NOV24',
            'prompt2': 'DEC24',
            'base_mid': 103.0,
            'base_bid': 102.5,
            'base_ask': 103.5
        }
    ]
    
    # Create snapshots over the last 24 hours
    for hours_ago in range(24):
        for minutes in range(0, 60, 5):  # Every 5 minutes
            timestamp = now - timedelta(hours=hours_ago, minutes=minutes)
            
            # Create snapshots for each spread
            for spread in spreads:
                # Add some random-like price movement
                movement = (hours_ago + minutes/60) * 0.01
                
                snapshot = Snapshot(
                    timestamp=timestamp,
                    spread_name=spread['name'],
                    prompt1=spread['prompt1'],
                    prompt2=spread['prompt2'],
                    old_midpoint=spread['base_mid'] + movement,
                    new_midpoint=spread['base_mid'] + movement + 0.02,
                    old_bid=spread['base_bid'] + movement,
                    new_bid=spread['base_bid'] + movement + 0.02,
                    old_ask=spread['base_ask'] + movement,
                    new_ask=spread['base_ask'] + movement + 0.02
                )
                session.add(snapshot)
    
    # Add some larger price moves for testing
    big_moves = [
        (0.25, 'JUL24-AUG24'),
        (0.15, 'SEP24-OCT24'),
        (0.20, 'NOV24-DEC24')
    ]
    
    for move, spread_name in big_moves:
        spread = next(s for s in spreads if s['name'] == spread_name)
        snapshot = Snapshot(
            timestamp=now - timedelta(hours=2),
            spread_name=spread_name,
            prompt1=spread['prompt1'],
            prompt2=spread['prompt2'],
            old_midpoint=spread['base_mid'],
            new_midpoint=spread['base_mid'] + move,
            old_bid=spread['base_bid'],
            new_bid=spread['base_bid'] + move,
            old_ask=spread['base_ask'],
            new_ask=spread['base_ask'] + move
        )
        session.add(snapshot)
    
    session.commit()
    session.close()
    
    print("Mock data populated successfully!")

if __name__ == '__main__':
    populate_mock_data() 