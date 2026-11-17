from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.draft_deletion_preview import DraftDeletionPreview


class DraftDeletionPreviewView(BaseModel):
    """What preview_draft_deletion hands back: enough to confirm, and the handle to proceed."""

    model_config = ConfigDict(frozen=True)

    token: str = Field(
        description="Opaque, single-use handle for delete_draft. Valid in this session only."
    )
    summary: str = Field(description="One line identifying the draft that would be deleted.")

    @classmethod
    def from_preview(cls, preview: DraftDeletionPreview) -> "DraftDeletionPreviewView":
        return cls(token=preview.token, summary=preview.summary)
