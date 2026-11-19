from pydantic import BaseModel, ConfigDict, Field

MOVED_NOTICE = (
    "The draft was moved to Deleted Items. It can still be recovered from there in Outlook."
)


class DraftDeletionResultView(BaseModel):
    """What delete_draft hands back, with where the draft went stated outright."""

    model_config = ConfigDict(frozen=True)

    id: str = Field(description="Graph id the deleted draft had.")
    notice: str = Field(description="Where the draft went and how to get it back.")

    @classmethod
    def from_draft_id(cls, draft_id: str) -> "DraftDeletionResultView":
        return cls(id=draft_id, notice=MOVED_NOTICE)
