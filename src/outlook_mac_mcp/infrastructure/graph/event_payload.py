from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_event import NewEvent

REQUIRED_ATTENDEE = "required"


def to_event_payload(new_event: NewEvent, timezone: ZoneInfo) -> dict[str, Any]:
    """Shape a NewEvent the way POST /me/events expects it.

    Graph takes a wall clock plus a zone name, never an offset, so a timed event is
    converted into the resolved zone and sent with that zone's name. An all-day event is
    not converted: its date is what the user meant, and a conversion could move midnight
    into the previous day.
    """
    payload: dict[str, Any] = {
        "subject": new_event.subject,
        "start": _date_time(new_event, new_event.start, timezone),
        "end": _date_time(new_event, new_event.end, timezone),
        "isAllDay": new_event.is_all_day,
        "attendees": [_attendee(attendee) for attendee in new_event.attendees],
    }
    if new_event.location:
        payload["location"] = {"displayName": new_event.location}
    return payload


def _date_time(new_event: NewEvent, instant: datetime, timezone: ZoneInfo) -> dict[str, str]:
    wall_clock = (
        instant.replace(tzinfo=None) if new_event.is_all_day else instant.astimezone(timezone)
    )
    return {"dateTime": wall_clock.replace(tzinfo=None).isoformat(), "timeZone": timezone.key}


def _attendee(attendee: EmailAddress) -> dict[str, Any]:
    email_address: dict[str, str] = {"address": attendee.address}
    if attendee.display_name:
        email_address["name"] = attendee.display_name
    return {"emailAddress": email_address, "type": REQUIRED_ATTENDEE}
