from datetime import datetime


def eastern_now(timezone_factory=None, now_factory=None, warning=None):
    """Return (current time, zone_available), preserving 0DTE on fallback."""
    timezone_factory = timezone_factory or _eastern_timezone
    now_factory = now_factory or datetime.now
    try:
        timezone = timezone_factory()
        return now_factory(timezone), True
    except (ImportError, KeyError, OSError, RuntimeError, ValueError) as exc:
        if warning is not None:
            warning(exc)
        return now_factory(), False


def _eastern_timezone():
    import pytz

    return pytz.timezone("America/New_York")


def filter_expiry_dates(all_expiries, now, eastern_available):
    today = now.strftime("%Y-%m-%d")
    dates = sorted(expiry for expiry in all_expiries if expiry >= today)
    if (
        eastern_available
        and now.hour >= 16
        and dates
        and dates[0] == today
    ):
        dates = dates[1:]
    return dates
