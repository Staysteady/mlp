from market_maker.utils.db_monitor import DatabaseMonitor
from functools import wraps


def with_monitor(func):
    """Decorator that provides a DatabaseMonitor to the decorated function."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        with DatabaseMonitor() as monitor:
            return func(monitor, *args, **kwargs)
    return wrapper 