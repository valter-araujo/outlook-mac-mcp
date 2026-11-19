from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.mail_draft_preview import MailDraftPreview


class MailDraftPreviewView(BaseModel):
    """What preview_draft hands back: enough to confirm, and the handle to proceed."""

    model_config = ConfigDict(frozen=True)

    token: str = Field(
        description="Opaque, single-use handle for create_draft. Valid in this session only."
    )
    summary: str = Field(description="One line describing the draft exactly as it would be saved.")

    @classmethod
    def from_preview(cls, preview: MailDraftPreview) -> "MailDraftPreviewView":
        return cls(token=preview.token, summary=preview.summary)
