from typing import Protocol

from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent


class CalendarWriter(Protocol):
    def create(self, new_event: NewEvent) -> Event:
        """Create `new_event` on the default calendar and return it as stored."""
        ...

    def get_by_id(self, event_id: str) -> Event:
        """Return the event with `event_id`, body and attendees included.

        Raises EventNotFoundError when the calendar holds no such event. Used to build
        an update diff or a deletion snapshot before either is previewed, never by a
        listing.
        """
        ...

    def update(self, changes: EventChanges) -> Event:
        """Apply `changes` to the event it names and return it as stored, with only the
        fields `changes` actually carries sent to Graph.

        Raises EventNotFoundError when the calendar holds no such event.
        """
        ...

    def delete(self, event_id: str) -> None:
        """Remove the event with `event_id` from the default calendar. Irreversible.

        Raises EventNotFoundError when the calendar holds no such event.
        """
        ...
