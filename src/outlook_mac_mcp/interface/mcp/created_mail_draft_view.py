from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft

NOT_SENT_NOTICE = (
    "The draft is saved in the Drafts folder and has NOT been sent. This server cannot "
    "send mail; the user reviews and sends it from Outlook."
)


class CreatedMailDraftView(BaseModel):
    """What create_draft hands back, with the fact that matters most stated outright."""

    model_config = ConfigDict(frozen=True)

    id: str = Field(description="Graph id of the saved draft.")
    notice: str = Field(description="Reminder that the draft was saved, not sent.")

    @classmethod
    def from_created(cls, created: CreatedMailDraft) -> "CreatedMailDraftView":
        return cls(id=created.id, notice=NOT_SENT_NOTICE)
