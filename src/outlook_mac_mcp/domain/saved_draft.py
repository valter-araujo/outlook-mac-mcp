from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class SavedDraft:
    """An unsent message confirmed to be in the Drafts folder, as a deletion shows it."""

    id: str
    subject: str
    to: tuple[EmailAddress, ...]
    last_modified_at: datetime
