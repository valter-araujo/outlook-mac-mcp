from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.new_event import NewEvent

ORGANIZER = EmailAddress(address="me@example.com", display_name="Me")


class InMemoryCalendarWriter:
    """Fake CalendarWriter that remembers what it was asked to create."""

    def __init__(self) -> None:
        self.created: list[NewEvent] = []

    def create(self, new_event: NewEvent) -> Event:
        self.created.append(new_event)
        return Event(
            id=f"created-{len(self.created)}",
            subject=new_event.subject,
            start=new_event.start,
            end=new_event.end,
            is_all_day=new_event.is_all_day,
            location=new_event.location,
            organizer=ORGANIZER,
        )
