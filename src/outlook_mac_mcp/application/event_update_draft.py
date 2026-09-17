from dataclasses import dataclass

from outlook_mac_mcp.domain.event_changes import EventChanges


@dataclass(frozen=True, slots=True)
class EventUpdateDraft:
    """A validated update waiting for the user's go-ahead.

    Carries `changes` rather than a merged event: UpdateEvent sends a PATCH with only
    the fields that were actually asked to change, and `changes` already holds exactly
    those, with `None` for everything left alone.
    """

    token: str
    summary: str
    changes: EventChanges
