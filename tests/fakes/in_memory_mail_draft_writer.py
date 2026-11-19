from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft


class InMemoryMailDraftWriter:
    """Fake MailDraftWriter that remembers what it was asked to save."""

    def __init__(self) -> None:
        self.created: list[NewMailDraft] = []

    def create(self, new_draft: NewMailDraft) -> CreatedMailDraft:
        self.created.append(new_draft)
        return CreatedMailDraft(id=f"draft-{len(self.created)}")
