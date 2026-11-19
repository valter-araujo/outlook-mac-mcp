from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.errors import EmailNotFoundError
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_source import ReplySource


class InMemoryReplyDraftWriter:
    """Fake ReplyDraftWriter over a fixed set of messages, remembering what it saved."""

    def __init__(
        self,
        sources: tuple[ReplySource, ...] = (),
        own_addresses: frozenset[str] = frozenset(),
    ) -> None:
        self._sources = {source.id: source for source in sources}
        self._own_addresses = own_addresses
        self.created: list[NewReplyDraft] = []

    def get_reply_source(self, email_id: str) -> ReplySource:
        try:
            return self._sources[email_id]
        except KeyError:
            raise EmailNotFoundError(f"no email with id {email_id}") from None

    def get_own_addresses(self) -> frozenset[str]:
        return self._own_addresses

    def create_reply(self, draft: NewReplyDraft) -> CreatedMailDraft:
        self.created.append(draft)
        return CreatedMailDraft(id=f"reply-{len(self.created)}")
