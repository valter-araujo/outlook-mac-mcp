from collections.abc import Mapping
from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import to_email_address
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import (
    optional_text,
    optional_text_body,
    required_flag,
    required_text,
)

EVENT_FIELDS = ("id", "subject", "start", "end", "isAllDay", "location", "organizer")


def to_event(event: Mapping[str, Any]) -> Event:
    """Turn one Graph event resource into a domain Event.

    Graph sends each time as a wall clock plus a zone name, never as an offset. The zone
    is the one asked for in the Prefer header, except that all-day events may come back
    in the zone they were created in; attaching whatever zone Graph names keeps the
    instant right in both cases.

    A listing never selects `body` or `attendees`, so both read as empty there; only a
    targeted read of one event, with those fields selected, ever carries them.
    """
    return Event(
        id=required_text(event, "id"),
        subject=optional_text(event, "subject"),
        start=_read_date_time(event, "start"),
        end=_read_date_time(event, "end"),
        is_all_day=required_flag(event, "isAllDay"),
        location=_read_location(event),
        organizer=to_email_address(event.get("organizer")),
        body=optional_text_body(event),
        attendees=_read_attendees(event),
    )


def _read_attendees(event: Mapping[str, Any]) -> tuple[EmailAddress, ...]:
    attendees = event.get("attendees")
    if not isinstance(attendees, list):
        return ()
    return tuple(to_email_address(attendee) for attendee in attendees)


def _read_date_time(event: Mapping[str, Any], field: str) -> datetime:
    value = event.get(field)
    if not isinstance(value, dict):
        raise GraphResponseError(f"the event carried no {field}")
    zone = _read_zone(value, field)
    raw = required_text(value, "dateTime")
    try:
        wall_clock = datetime.fromisoformat(raw)
    except ValueError as error:
        raise GraphResponseError(f"{field}.dateTime was not an ISO 8601 timestamp") from error
    if wall_clock.tzinfo is not None:
        return wall_clock.astimezone(zone)
    return wall_clock.replace(tzinfo=zone)


def _read_zone(date_time: Mapping[str, Any], field: str) -> ZoneInfo:
    name = required_text(date_time, "timeZone")
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as error:
        raise GraphResponseError(f"{field}.timeZone {name!r} is not an IANA time zone") from error


def _read_location(event: Mapping[str, Any]) -> str:
    location = event.get("location")
    if not isinstance(location, dict):
        return ""
    return optional_text(location, "displayName")
