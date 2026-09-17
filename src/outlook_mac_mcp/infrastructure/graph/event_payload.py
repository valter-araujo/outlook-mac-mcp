from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event_changes import EventChanges
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
        "start": _date_time(new_event.is_all_day, new_event.start, timezone),
        "end": _date_time(new_event.is_all_day, new_event.end, timezone),
        "isAllDay": new_event.is_all_day,
        "attendees": [_attendee(attendee) for attendee in new_event.attendees],
    }
    if new_event.location:
        payload["location"] = {"displayName": new_event.location}
    if new_event.body:
        payload["body"] = {"contentType": "text", "content": new_event.body}
    if new_event.reminder_minutes_before_start is not None:
        payload["reminderMinutesBeforeStart"] = new_event.reminder_minutes_before_start
    if new_event.sensitivity is not None:
        payload["sensitivity"] = new_event.sensitivity.value
    if new_event.show_as is not None:
        payload["showAs"] = new_event.show_as.value
    return payload


def to_event_patch_payload(
    changes: EventChanges, is_all_day: bool, timezone: ZoneInfo
) -> dict[str, Any]:
    """Shape an EventChanges the way PATCH /me/events/{id} expects it: only the fields
    the caller actually supplied, so an untouched field is left exactly as it was.
    """
    payload: dict[str, Any] = {}
    if changes.subject is not None:
        payload["subject"] = changes.subject
    if changes.start is not None:
        payload["start"] = _date_time(is_all_day, changes.start, timezone)
    if changes.end is not None:
        payload["end"] = _date_time(is_all_day, changes.end, timezone)
    if changes.location is not None:
        payload["location"] = {"displayName": changes.location}
    if changes.body is not None:
        payload["body"] = {"contentType": "text", "content": changes.body}
    if changes.attendees is not None:
        payload["attendees"] = [_attendee(attendee) for attendee in changes.attendees]
    if changes.reminder_minutes_before_start is not None:
        payload["reminderMinutesBeforeStart"] = changes.reminder_minutes_before_start
    if changes.sensitivity is not None:
        payload["sensitivity"] = changes.sensitivity.value
    if changes.show_as is not None:
        payload["showAs"] = changes.show_as.value
    return payload


def _date_time(is_all_day: bool, instant: datetime, timezone: ZoneInfo) -> dict[str, str]:
    wall_clock = instant.replace(tzinfo=None) if is_all_day else instant.astimezone(timezone)
    return {"dateTime": wall_clock.replace(tzinfo=None).isoformat(), "timeZone": timezone.key}


def _attendee(attendee: EmailAddress) -> dict[str, Any]:
    email_address: dict[str, str] = {"address": attendee.address}
    if attendee.display_name:
        email_address["name"] = attendee.display_name
    return {"emailAddress": email_address, "type": REQUIRED_ATTENDEE}
