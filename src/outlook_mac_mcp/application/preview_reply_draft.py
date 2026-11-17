from dataclasses import dataclass

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.get_email import MAX_EMAIL_ID_LENGTH
from outlook_mac_mcp.application.ports.reply_draft_writer import ReplyDraftWriter
from outlook_mac_mcp.application.reply_draft_preview import ReplyDraftPreview
from outlook_mac_mcp.application.reply_draft_summary import (
    describe_reply_draft,
    reply_draft_warnings,
)
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import resolve_reply_recipients


@dataclass(frozen=True, slots=True)
class ReplyDraftRequest:
    email_id: str
    mode: ReplyMode
    body: str
    extra_to: tuple[EmailAddress, ...] = ()
    extra_cc: tuple[EmailAddress, ...] = ()

    def __post_init__(self) -> None:
        if not self.email_id:
            raise InvalidRequestError("email_id must not be empty")
        if len(self.email_id) > MAX_EMAIL_ID_LENGTH:
            raise InvalidRequestError(f"email_id must be at most {MAX_EMAIL_ID_LENGTH} characters")


class PreviewReplyDraft:
    """First half of a reply: reads the original and the account, writes nothing.

    The recipients are resolved here, not left to Graph, so what the user confirms is
    the exact list CreateReplyDraft will save.
    """

    def __init__(
        self, reply_draft_writer: ReplyDraftWriter, draft_store: DraftStore[ReplyDraftPreview]
    ) -> None:
        self._reply_draft_writer = reply_draft_writer
        self._draft_store = draft_store

    def execute(self, request: ReplyDraftRequest) -> ReplyDraftPreview:
        source = self._reply_draft_writer.get_reply_source(request.email_id)
        own_addresses = self._reply_draft_writer.get_own_addresses()
        recipients = resolve_reply_recipients(source, request.mode, own_addresses)
        draft = NewReplyDraft(
            source_id=source.id,
            mode=request.mode,
            body=request.body,
            recipients=recipients.with_extra(request.extra_to, request.extra_cc),
        )
        summary = describe_reply_draft(source, draft)
        warnings = reply_draft_warnings(source, draft)
        return self._draft_store.add(
            lambda token: ReplyDraftPreview(
                token=token, summary=summary, warnings=warnings, draft=draft
            )
        )
