from outlook_mac_mcp.application.event_summary import describe_moment, summarize_body
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent


def merge_changes(current: Event, changes: EventChanges) -> NewEvent:
    """The event that applying `changes` to `current` would produce, validated by
    NewEvent's own constructor: the same rules a brand new event must satisfy, since a
    merged one must satisfy them too, and the one place those rules are written down.

    `is_all_day` is never a changeable field, so it always carries over from `current`;
    that is also what lets an all-day event's midnight-boundary rule apply correctly to
    a start- or end-only change, the same way it would to a full one.
    """
    return NewEvent(
        subject=changes.subject if changes.subject is not None else current.subject,
        start=changes.start if changes.start is not None else current.start,
        end=changes.end if changes.end is not None else current.end,
        is_all_day=current.is_all_day,
        location=changes.location if changes.location is not None else current.location,
        body=changes.body if changes.body is not None else current.body,
        attendees=changes.attendees if changes.attendees is not None else current.attendees,
    )


def describe_changes(current: Event, changes: EventChanges) -> str:
    """One line per changed field, old value -> new value, never just the resulting
    state: the preview -> token -> update contract only holds if the confirmation shows
    exactly what will change, not what the event will look like afterward.

    Leads with the current subject and time so the event being changed is identified
    even when neither is itself one of the changed fields — the same disambiguation a
    deletion preview needs when two similar events exist.
    """
    diffs: list[str] = []
    if changes.subject is not None:
        diffs.append(f'subject: "{current.subject}" -> "{changes.subject}"')
    if changes.start is not None:
        diffs.append(f"start: {describe_moment(current.start)} -> {describe_moment(changes.start)}")
    if changes.end is not None:
        diffs.append(f"end: {describe_moment(current.end)} -> {describe_moment(changes.end)}")
    if changes.location is not None:
        diffs.append(f"location: {_place(current.location)} -> {_place(changes.location)}")
    if changes.body is not None:
        diffs.append(f"body: {_body(current.body)} -> {_body(changes.body)}")
    if changes.attendees is not None:
        diffs.append(f"attendees: {_people(current.attendees)} -> {_people(changes.attendees)}")
    identity = (
        f'"{current.subject}" ({describe_moment(current.start)} to {describe_moment(current.end)})'
    )
    return f"{identity}: " + "; ".join(diffs)


def _place(location: str) -> str:
    return f'"{location}"' if location else "(none)"


def _body(body: str) -> str:
    return summarize_body(body) if body else "(none)"


def _people(attendees: tuple[EmailAddress, ...]) -> str:
    return ", ".join(attendee.address for attendee in attendees) if attendees else "(none)"
