from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.event_deletion_draft import EventDeletionDraft


class EventDeletionDraftView(BaseModel):
    """What preview_event_deletion hands back: enough to confirm, and the handle to proceed."""

    model_config = ConfigDict(frozen=True)

    token: str = Field(
        description="Opaque, single-use handle for delete_event. Valid in this session only."
    )
    summary: str = Field(
        description="Every field of the event as it currently stands, unabbreviated: "
        "exactly what will be removed."
    )

    @classmethod
    def from_draft(cls, draft: EventDeletionDraft) -> "EventDeletionDraftView":
        return cls(token=draft.token, summary=draft.summary)
