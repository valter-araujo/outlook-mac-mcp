from dataclasses import dataclass

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class ReplySource:
    """The parts of an existing message that decide who a reply to it goes to.

    `reply_to` is the message's Reply-To header. When present it replaces the sender as
    the reply's addressee, which is what Outlook and Graph's createReply both do and what
    applicant-tracking systems rely on when they send from one address and read another.
    """

    id: str
    subject: str
    sender: EmailAddress
    reply_to: tuple[EmailAddress, ...]
    to: tuple[EmailAddress, ...]
    cc: tuple[EmailAddress, ...]
