from datetime import datetime, timedelta

from outlook_mac_mcp.domain.email_address import EmailAddress
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
        parts.append(f"body: {summarize_body(event.body)}")
    return "; ".join(parts)


def summarize_body(body: str) -> str:
    """The body as it will appear in a one-line summary: quoted and verbatim if it fits
    a reasonable display length, or quoted, truncated and counted when it does not, so
    nothing that will be written or removed is ever silently hidden from a confirmation.
    """
    if len(body) <= MAX_BODY_IN_SUMMARY:
        return f'"{body}"'
    truncated = body[:MAX_BODY_IN_SUMMARY]
    return f'"{truncated}…" ({len(body)} characters total, truncated for this summary)'


def describe_moment(instant: datetime) -> str:
    """One point in time, for an update diff or a deletion snapshot, where `describe`'s
    range-and-all-day-aware formatting is not what is being shown, only two boundaries.
    """
    offset = f"UTC{instant.isoformat()[-6:]}"
    return f"{instant.strftime(DAY_FORMAT)} {instant.strftime(CLOCK_FORMAT)} ({offset})"


def describe_place(location: str) -> str:
    """A location field for an update diff or a deletion snapshot, where an empty value
    must still be shown explicitly rather than producing a blank or missing line.
    """
    return f'"{location}"' if location else "(none)"


def describe_body(body: str) -> str:
    """A body field for an update diff or a deletion snapshot; see describe_place."""
    return summarize_body(body) if body else "(none)"


def describe_attendees(attendees: tuple[EmailAddress, ...]) -> str:
    """An attendee list for an update diff or a deletion snapshot; see describe_place."""
    return ", ".join(attendee.address for attendee in attendees) if attendees else "(none)"


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
