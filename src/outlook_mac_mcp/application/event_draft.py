from dataclasses import dataclass

from outlook_mac_mcp.domain.new_event import NewEvent


@dataclass(frozen=True, slots=True)
class EventDraft:
    """A validated event waiting for the user's go-ahead.

    The token is the only handle a client gets back, so a create call can refer to
    exactly what was previewed and nothing else.
    """

    token: str
    summary: str
    new_event: NewEvent
