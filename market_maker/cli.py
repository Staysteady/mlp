"""
Command-line interface for the market maker system.
Provides easy access to database monitoring and log information.
"""
# Standard library imports
import click
import pandas as pd
from .utils.db_decorator import with_monitor
from .utils.logging_config import LOGS_DIR

# Set pandas display options for better CLI output
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

@click.group()
def cli():
    """Market Maker monitoring and management CLI."""

@cli.command()
@with_monitor
def stats(monitor):
    """Show general database statistics."""
    db_stats = monitor.get_database_stats()
    click.echo("\nDatabase Statistics:")
    click.echo("-------------------")
    for key, value in db_stats.items():
        click.echo(f"{key.replace('_', ' ').title()}: {value}")

@cli.command()
@click.option('--minutes', default=5, help='Number of minutes to look back')
@with_monitor
def recent(monitor, minutes):
    """Show recent snapshots."""
    snapshots = monitor.get_recent_snapshots(minutes=minutes)
    click.echo(f"\nRecent Snapshots (last {minutes} minutes):")
    click.echo("----------------------------------------")
    for snap in snapshots:
        click.echo(f"Time: {snap.timestamp}, Spread: {snap.spread_name}, "
                  f"Mid: {snap.old_midpoint} -> {snap.new_midpoint}")

@cli.command()
@click.argument('spread_name')
@click.option('--hours', default=24, help='Number of hours to look back')
@with_monitor
def history(monitor, spread_name, hours):
    """Show price history for a specific spread."""
    df = monitor.get_spread_history(spread_name, hours=hours)
    if df.empty:
        click.echo(f"No history found for spread {spread_name}")
        return
    click.echo(f"\nPrice History for {spread_name} (last {hours} hours):")
    click.echo("------------------------------------------------")
    click.echo(df.to_string())

@cli.command()
@click.option('--top-n', default=10, help='Number of largest moves to show')
@with_monitor
def moves(monitor, top_n):
    """Show largest price moves."""
    df = monitor.get_largest_moves(top_n=top_n)
    if df.empty:
        click.echo("No price moves found")
        return
    click.echo(f"\nTop {top_n} Largest Price Moves:")
    click.echo("---------------------------")
    click.echo(df.to_string())

@cli.command()
@click.option('--hours', default=24, help='Number of hours to look back')
@with_monitor
def summary(monitor, hours):
    """Show summary of all spreads activity."""
    df = monitor.get_spread_summary(hours=hours)
    if df.empty:
        click.echo("No spread activity found")
        return
    click.echo(f"\nSpread Activity Summary (last {hours} hours):")
    click.echo("----------------------------------------")
    click.echo(df.to_string())

@cli.command()
@click.option('--lines', default=50, help='Number of lines to show')
@click.option('--component', type=click.Choice(['main', 'database', 'excel']),
              default='main', help='Which component logs to show')
def logs(lines, component):
    """Show recent log entries."""
    log_files = {
        'main': LOGS_DIR / "market_maker.log",
        'database': LOGS_DIR / "database.log",
        'excel': LOGS_DIR / "excel_reader.log"
    }
    log_file = log_files[component]
    if not log_file.exists():
        click.echo(f"No log file found for {component}")
        return
    click.echo(f"\nRecent {component} logs (last {lines} lines):")
    click.echo("----------------------------------------")
    with open(log_file, encoding='utf-8') as f:
        for line in list(f)[-lines:]:
            click.echo(line.strip())

if __name__ == '__main__':
    cli() 