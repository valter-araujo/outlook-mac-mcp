from datetime import datetime, timedelta

from outlook_mac_mcp.domain.new_event import NewEvent

DAY_FORMAT = "%a %d %b %Y"
CLOCK_FORMAT = "%H:%M"


def describe(event: NewEvent) -> str:
    """One line a person can check against what they meant, before anything is created.

    The offset is spelled out because the client and the user may not share a zone, and
    an all-day range shows its last day inclusive, the way a calendar displays it.
    """
    parts = [event.subject, _when(event)]
    if event.location:
        parts.append(f"at {event.location}")
    if event.attendees:
        parts.append("with " + ", ".join(attendee.address for attendee in event.attendees))
    return "; ".join(parts)


def _when(event: NewEvent) -> str:
    if event.is_all_day:
        return _all_day_range(event.start, event.end)
    offset = f"UTC{event.start.isoformat()[-6:]}"
    if event.start.date() == event.end.date():
        return (
            f"{event.start.strftime(DAY_FORMAT)} {event.start.strftime(CLOCK_FORMAT)} to "
            f"{event.end.strftime(CLOCK_FORMAT)} ({offset})"
        )
    return (
        f"{event.start.strftime(f'{DAY_FORMAT} {CLOCK_FORMAT}')} to "
        f"{event.end.strftime(f'{DAY_FORMAT} {CLOCK_FORMAT}')} ({offset})"
    )


def _all_day_range(start: datetime, end: datetime) -> str:
    last_day = end - timedelta(days=1)
    if last_day.date() == start.date():
        return f"all day on {start.strftime(DAY_FORMAT)}"
    return f"all day from {start.strftime(DAY_FORMAT)} to {last_day.strftime(DAY_FORMAT)}"
