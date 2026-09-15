from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.event_mapper import to_event
from outlook_mac_mcp.infrastructure.graph.event_payload import to_event_payload
from outlook_mac_mcp.infrastructure.graph.preferences import timezone_preference

EVENTS_PATH = "/me/events"


class GraphCalendarWriter:
    """Creates events on the default calendar; the only code path that writes to Graph.

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
            timezone_preference(self._timezone),
        )
        return to_event(payload)
