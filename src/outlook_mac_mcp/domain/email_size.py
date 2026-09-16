from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class EmailSize:
    """One email as read for a size scan: enough to identify it, plus its size in bytes."""

    id: str
    subject: str
    sender: EmailAddress
    received_at: datetime
    size_bytes: int
