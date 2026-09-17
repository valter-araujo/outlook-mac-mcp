from http import HTTPStatus
from urllib.parse import quote
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.errors import EventNotFoundError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
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
EVENT_DETAIL_FIELDS = (*EVENT_FIELDS, "body", "attendees")


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
            self._create_headers(new_event),
        )
        return to_event(payload)

    def get_by_id(self, event_id: str) -> Event:
        try:
            payload = self._client.get(
                f"{EVENTS_PATH}/{quote(event_id, safe='')}",
                {"$select": ",".join(EVENT_DETAIL_FIELDS)},
                self._detail_headers(),
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
                self._detail_headers(),
            )
        except GraphRequestError as error:
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EventNotFoundError(f"no event with id {changes.event_id}") from error
            raise
        return to_event(payload)

    def delete(self, event_id: str) -> None:
        try:
            self._client.delete(f"{EVENTS_PATH}/{quote(event_id, safe='')}")
        except GraphRequestError as error:
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EventNotFoundError(f"no event with id {event_id}") from error
            raise

    def _resolve_is_all_day(self, changes: EventChanges) -> bool:
        """is_all_day is never itself a changeable field, so it is only worth a fetch when
        start or end is: those are the only payload fields it affects.
        """
        if changes.start is None and changes.end is None:
            return False
        return self.get_by_id(changes.event_id).is_all_day

    def _create_headers(self, new_event: NewEvent) -> dict[str, str]:
        """A text-body preference is only added when a body is actually being sent: with
        none, the response has nothing to be read back as HTML by default, and the
        request stays exactly what it was before this preference existed.
        """
        timezone = timezone_preference(self._timezone)
        if not new_event.body:
            return timezone
        return combined_preference(timezone, text_body_preference())

    def _detail_headers(self) -> dict[str, str]:
        """Unlike create, a text-body preference is always included: the event being read
        or patched may already carry a body regardless of what this call touches, and
        without the preference Graph would answer with HTML the mapper cannot read.
        """
        return combined_preference(timezone_preference(self._timezone), text_body_preference())
