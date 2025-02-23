#!/usr/bin/env python3
"""
Script to view the contents of the market maker database.
Provides easy-to-use functions to view recent changes, statistics, and spread history.
"""
from datetime import datetime, timedelta
import pandas as pd
from market_maker.data.models import Session, Snapshot
from market_maker.utils.db_decorator import with_monitor

@with_monitor
def show_stats(monitor):
    """Show general database statistics."""
    stats = monitor.get_database_stats()
    print("\nDatabase Statistics:")
    print("-------------------")
    for key, value in stats.items():
        print(f"{key.replace('_', ' ').title()}: {value}")

@with_monitor
def show_recent(monitor, minutes=5):
    """Show recent price changes."""
    snapshots = monitor.get_recent_snapshots(minutes=minutes)
    print(f"\nRecent Changes (last {minutes} minutes):")
    print("----------------------------------------")
    for snap in snapshots:
        change = snap.new_midpoint - snap.old_midpoint
        print(f"Time: {snap.timestamp}, Spread: {snap.spread_name:15}, "
              f"Mid: {snap.old_midpoint:7.2f} -> {snap.new_midpoint:7.2f} "
              f"(Δ: {change:+6.2f})")

@with_monitor
def show_spread_history(monitor, spread_name, hours=24):
    """Show price history for a specific spread."""
    df = monitor.get_spread_history(spread_name, hours=hours)
    if df.empty:
        print(f"No history found for spread {spread_name}")
        return
    print(f"\nPrice History for {spread_name} (last {hours} hours):")
    print("------------------------------------------------")
    print(df.to_string())

@with_monitor
def show_largest_moves(monitor, top_n=10):
    """Show largest price moves."""
    df = monitor.get_largest_moves(top_n=top_n)
    if df.empty:
        print("No price moves found")
        return
    print(f"\nTop {top_n} Largest Price Moves:")
    print("---------------------------")
    print(df.to_string())

@with_monitor
def show_spread_summary(monitor, hours=24):
    """Show summary of all spreads activity."""
    df = monitor.get_spread_summary(hours=hours)
    if df.empty:
        print("No spread activity found")
        return
    print(f"\nSpread Activity Summary (last {hours} hours):")
    print("----------------------------------------")
    print(df.to_string())

if __name__ == "__main__":
    import sys
    
    # Set pandas display options for better output
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    
    if len(sys.argv) < 2:
        print("\nUsage:")
        print("  python view_database.py stats              - Show database statistics")
        print("  python view_database.py recent [minutes]   - Show recent changes")
        print("  python view_database.py history SPREAD [hours] - Show spread history")
        print("  python view_database.py moves [top_n]      - Show largest moves")
        print("  python view_database.py summary [hours]    - Show spread summary")
        sys.exit(1)
    
    command = sys.argv[1]
    
    if command == "stats":
        show_stats()
    elif command == "recent":
        minutes = int(sys.argv[2]) if len(sys.argv) > 2 else 5
        show_recent(minutes)
    elif command == "history":
        if len(sys.argv) < 3:
            print("Error: Please specify a spread name")
            sys.exit(1)
        spread = sys.argv[2]
        hours = int(sys.argv[3]) if len(sys.argv) > 3 else 24
        show_spread_history(spread, hours)
    elif command == "moves":
        top_n = int(sys.argv[2]) if len(sys.argv) > 2 else 10
        show_largest_moves(top_n)
    elif command == "summary":
        hours = int(sys.argv[2]) if len(sys.argv) > 2 else 24
        show_spread_summary(hours)
    else:
        print(f"Unknown command: {command}")
        sys.exit(1) 