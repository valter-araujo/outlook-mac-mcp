from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.event_update_draft import EventUpdateDraft


class EventUpdateDraftView(BaseModel):
    """What preview_event_update hands back: enough to confirm, and the handle to proceed."""

    model_config = ConfigDict(frozen=True)

    token: str = Field(
        description="Opaque, single-use handle for update_event. Valid in this session only."
    )
    summary: str = Field(
        description="The event being changed, followed by one clause per changed field "
        "as old value -> new value. Fields left out of the call do not appear here and "
        "will not change."
    )

    @classmethod
    def from_draft(cls, draft: EventUpdateDraft) -> "EventUpdateDraftView":
        return cls(token=draft.token, summary=draft.summary)
