from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.reply_draft_preview import ReplyDraftPreview


class ReplyDraftPreviewView(BaseModel):
    """What preview_reply_draft hands back: enough to confirm, and the handle to proceed."""

    model_config = ConfigDict(frozen=True)

    token: str = Field(
        description="Opaque, single-use handle for create_reply_draft. Valid in this session only."
    )
    summary: str = Field(
        description="One line describing the reply exactly as it would be saved, every "
        "recipient included."
    )
    to: tuple[str, ...] = Field(description="The final To line, as it will be saved.")
    cc: tuple[str, ...] = Field(description="The final Cc line, as it will be saved.")
    warnings: tuple[str, ...] = Field(
        description="Things the user should know before agreeing, such as a no-reply "
        "recipient. Show every one of them to the user."
    )

    @classmethod
    def from_preview(cls, preview: ReplyDraftPreview) -> "ReplyDraftPreviewView":
        recipients = preview.draft.recipients
        return cls(
            token=preview.token,
            summary=preview.summary,
            to=tuple(recipient.address for recipient in recipients.to),
            cc=tuple(recipient.address for recipient in recipients.cc),
            warnings=preview.warnings,
        )
