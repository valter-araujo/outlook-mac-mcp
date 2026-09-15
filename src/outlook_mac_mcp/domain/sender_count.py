from dataclasses import dataclass

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class SenderCount:
    sender: EmailAddress
    count: int
