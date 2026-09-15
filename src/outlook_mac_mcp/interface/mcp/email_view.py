from datetime import datetime

from pydantic import BaseModel, ConfigDict

from outlook_mac_mcp.domain.email import Email


class EmailView(BaseModel):
    """JSON-serializable projection of an Email.

    Field for field, with the sender flattened because JSON consumers read a flat
    record more easily than a nested one. Nothing is computed, filtered or reordered.
    """

    model_config = ConfigDict(frozen=True)

    id: str
    subject: str
    sender_address: str
    sender_name: str
    received_at: datetime
    is_read: bool
    has_attachments: bool
    preview: str

    @classmethod
    def from_email(cls, email: Email) -> "EmailView":
        return cls(
            id=email.id,
            subject=email.subject,
            sender_address=email.sender.address,
            sender_name=email.sender.display_name,
            received_at=email.received_at,
            is_read=email.is_read,
            has_attachments=email.has_attachments,
            preview=email.preview,
        )
