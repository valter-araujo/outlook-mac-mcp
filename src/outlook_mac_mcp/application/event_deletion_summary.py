from outlook_mac_mcp.application.event_summary import (
    describe_attendees,
    describe_body,
    describe_moment,
    describe_place,
)
from outlook_mac_mcp.domain.event import Event


def describe_event_for_deletion(event: Event) -> str:
    """The full current event, every field shown, none abbreviated away: the
    confirmation must make plain exactly what is about to be removed, which matters most
    when two similar events exist and only their full detail tells them apart.
    """
    fields = [
        f'subject: "{event.subject}"',
        f"start: {describe_moment(event.start)}",
        f"end: {describe_moment(event.end)}",
        f"location: {describe_place(event.location)}",
        f"body: {describe_body(event.body)}",
        f"attendees: {describe_attendees(event.attendees)}",
    ]
    return "; ".join(fields)
