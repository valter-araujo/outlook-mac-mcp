from datetime import datetime, timedelta

from outlook_mac_mcp.domain.new_event import NewEvent

DAY_FORMAT = "%a %d %b %Y"
CLOCK_FORMAT = "%H:%M"
# How much of the body appears verbatim in the one-line summary before it is truncated.
# The preview -> token -> create contract only holds if the preview shows everything
# that will be written, so a body past this length is still shown in full, just with an
# explicit character count in place of the part that does not fit on one line.
MAX_BODY_IN_SUMMARY = 500


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
    if event.body:
        parts.append(_body_summary(event.body))
    return "; ".join(parts)


def _body_summary(body: str) -> str:
    if len(body) <= MAX_BODY_IN_SUMMARY:
        return f'body: "{body}"'
    truncated = body[:MAX_BODY_IN_SUMMARY]
    return f'body: "{truncated}…" ({len(body)} characters total, truncated for this summary)'


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
