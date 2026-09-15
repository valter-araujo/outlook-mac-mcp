from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.event_draft import EventDraft


class EventDraftView(BaseModel):
    """What preview_event hands back: enough to confirm, and the handle to proceed."""

    model_config = ConfigDict(frozen=True)

    token: str = Field(
        description="Opaque, single-use handle for create_event. Valid in this session only."
    )
    summary: str = Field(
        description="One line describing the event exactly as it would be created."
    )

    @classmethod
    def from_draft(cls, draft: EventDraft) -> "EventDraftView":
        return cls(token=draft.token, summary=draft.summary)
