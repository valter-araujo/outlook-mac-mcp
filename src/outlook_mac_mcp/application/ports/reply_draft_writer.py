from typing import Protocol

from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_source import ReplySource


class ReplyDraftWriter(Protocol):
    def get_reply_source(self, email_id: str) -> ReplySource:
        """Return the sender and recipients of the message with `email_id`.

        Raises EmailNotFoundError when the mailbox holds no such message.
        """
        ...

    def get_own_addresses(self) -> frozenset[str]:
        """Return the signed-in account's addresses, keyed with address_key.

        Aliases are not included: Graph does not list them for a personal account.
        """
        ...

    def create_reply(self, draft: NewReplyDraft) -> CreatedMailDraft:
        """Save `draft` as a reply in the original message's conversation, in Drafts.

        Saves only: nothing is sent. Raises EmailNotFoundError when the original is gone.
        """
        ...
