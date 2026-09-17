from collections.abc import Mapping
from http import HTTPStatus
from typing import Any
from urllib.parse import quote
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.errors import EventNotFoundError, EventWriteUnconfirmedError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError
from outlook_mac_mcp.infrastructure.graph.event_mapper import EVENT_FIELDS, to_event
from outlook_mac_mcp.infrastructure.graph.event_payload import (
    to_event_patch_payload,
    to_event_payload,
)
from outlook_mac_mcp.infrastructure.graph.preferences import (
    combined_preference,
    text_body_preference,
    timezone_preference,
)

EVENTS_PATH = "/me/events"
EVENT_DETAIL_FIELDS = (
    *EVENT_FIELDS,
    "body",
    "attendees",
    "reminderMinutesBeforeStart",
    "sensitivity",
    "showAs",
)


class GraphCalendarWriter:
    """Reads and writes events on the default calendar; the only code path that writes to Graph.

    Satisfies the CalendarWriter port structurally. It is wired only when the write flag
    is on, which is also the only time the token carries the scope it needs.
    """

    def __init__(self, client: GraphClient, timezone: ZoneInfo) -> None:
        self._client = client
        self._timezone = timezone

    def create(self, new_event: NewEvent) -> Event:
        payload = self._client.post(
            EVENTS_PATH,
            to_event_payload(new_event, self._timezone),
            self._body_read_headers(),
        )
        return self._map_write_response(payload, "created")

    def get_by_id(self, event_id: str) -> Event:
        try:
            payload = self._client.get(
                f"{EVENTS_PATH}/{quote(event_id, safe='')}",
                {"$select": ",".join(EVENT_DETAIL_FIELDS)},
                self._body_read_headers(),
            )
        except GraphRequestError as error:
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EventNotFoundError(f"no event with id {event_id}") from error
            raise
        return to_event(payload)

    def update(self, changes: EventChanges) -> Event:
        is_all_day = self._resolve_is_all_day(changes)
        try:
            payload = self._client.patch(
                f"{EVENTS_PATH}/{quote(changes.event_id, safe='')}",
                to_event_patch_payload(changes, is_all_day, self._timezone),
                self._body_read_headers(),
            )
        except GraphRequestError as error:
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EventNotFoundError(f"no event with id {changes.event_id}") from error
            raise
        return self._map_write_response(payload, "updated")

    def delete(self, event_id: str) -> None:
        try:
            self._client.delete(f"{EVENTS_PATH}/{quote(event_id, safe='')}")
        except GraphRequestError as error:
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EventNotFoundError(f"no event with id {event_id}") from error
            raise

    def _map_write_response(self, payload: Mapping[str, Any], verb: str) -> Event:
        """By the time this runs, Graph has already created or updated the event: a
        mapping failure here must never be indistinguishable from the write itself
        failing, or a caller who retries turns one write into two.
        """
        try:
            return to_event(payload)
        except GraphResponseError as error:
            event_id = payload.get("id")
            where = f" (id {event_id})" if isinstance(event_id, str) else ""
            raise EventWriteUnconfirmedError(
                f"the event was {verb}{where}, but its details could not be read back "
                f"afterward: {error}. Check your calendar directly before retrying."
            ) from error

    def _resolve_is_all_day(self, changes: EventChanges) -> bool:
        """is_all_day is never itself a changeable field, so it is only worth a fetch when
        start or end is: those are the only payload fields it affects.
        """
        if changes.start is None and changes.end is None:
            return False
        return self.get_by_id(changes.event_id).is_all_day

    def _body_read_headers(self) -> dict[str, str]:
        """Always ask for the body as text, on create as much as on get_by_id and update.

        Create used to skip this when no body was being sent, on the assumption that
        Graph would then have "nothing to read back as HTML". That assumption was wrong:
        confirmed live, an event created with no body at all still comes back with a
        `body` property, and without this preference Graph defaults it to HTML — an
        Exchange-generated empty wrapper, not the absence of a body — which the mapper
        then refuses to read as text. Graph always returns a body, so the preference is
        never conditional on whether one was sent.
        """
        return combined_preference(timezone_preference(self._timezone), text_body_preference())
