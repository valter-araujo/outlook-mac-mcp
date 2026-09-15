from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class Email:
    id: str
    subject: str
    sender: EmailAddress
    received_at: datetime
    is_read: bool
    has_attachments: bool
    preview: str
