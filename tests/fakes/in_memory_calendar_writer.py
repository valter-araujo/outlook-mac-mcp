from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import EventNotFoundError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent

ORGANIZER = EmailAddress(address="me@example.com", display_name="Me")


class InMemoryCalendarWriter:
    """Fake CalendarWriter that remembers what it was asked to create or update, and
    holds pre-seeded events for get_by_id, update and delete to find.
    """

    def __init__(self) -> None:
        self.created: list[NewEvent] = []
        self.updated: list[EventChanges] = []
        self.deleted: list[str] = []
        self._events: dict[str, Event] = {}

    def seed(self, event: Event) -> None:
        self._events[event.id] = event

    def create(self, new_event: NewEvent) -> Event:
        self.created.append(new_event)
        event = Event(
            id=f"created-{len(self.created)}",
            subject=new_event.subject,
            start=new_event.start,
            end=new_event.end,
            is_all_day=new_event.is_all_day,
            location=new_event.location,
            organizer=ORGANIZER,
            body=new_event.body,
            attendees=new_event.attendees,
            reminder_minutes_before_start=new_event.reminder_minutes_before_start,
            sensitivity=new_event.sensitivity,
            show_as=new_event.show_as,
        )
        self._events[event.id] = event
        return event

    def get_by_id(self, event_id: str) -> Event:
        try:
            return self._events[event_id]
        except KeyError:
            raise EventNotFoundError(f"no event with id {event_id}") from None

    def update(self, changes: EventChanges) -> Event:
        self.updated.append(changes)
        current = self.get_by_id(changes.event_id)
        updated = Event(
            id=current.id,
            subject=changes.subject if changes.subject is not None else current.subject,
            start=changes.start if changes.start is not None else current.start,
            end=changes.end if changes.end is not None else current.end,
            is_all_day=current.is_all_day,
            location=changes.location if changes.location is not None else current.location,
            organizer=current.organizer,
            body=changes.body if changes.body is not None else current.body,
            attendees=changes.attendees if changes.attendees is not None else current.attendees,
            reminder_minutes_before_start=(
                changes.reminder_minutes_before_start
                if changes.reminder_minutes_before_start is not None
                else current.reminder_minutes_before_start
            ),
            sensitivity=(
                changes.sensitivity if changes.sensitivity is not None else current.sensitivity
            ),
            show_as=changes.show_as if changes.show_as is not None else current.show_as,
        )
        self._events[updated.id] = updated
        return updated

    def delete(self, event_id: str) -> None:
        self.get_by_id(event_id)
        self.deleted.append(event_id)
        del self._events[event_id]
