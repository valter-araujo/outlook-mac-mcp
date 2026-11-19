from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_summary import MAX_BODY_IN_SUMMARY
from outlook_mac_mcp.application.mail_draft_preview import MailDraftPreview
from outlook_mac_mcp.application.preview_mail_draft import PreviewMailDraft
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft

ANA = EmailAddress(address="ana@example.com")
BO = EmailAddress(address="bo@example.com")


def preview(draft: NewMailDraft) -> MailDraftPreview:
    return PreviewMailDraft(DraftStore()).execute(draft)


def test_parks_the_draft_in_the_store_under_the_returned_token() -> None:
    store: DraftStore[MailDraftPreview] = DraftStore()
    draft = NewMailDraft(subject="Hello", body="Hi.", to=(ANA,))

    result = PreviewMailDraft(store).execute(draft)

    assert store.take(result.token).new_draft == draft


def test_describes_subject_recipients_and_body() -> None:
    draft = NewMailDraft(subject="Hello", body="Hi Ana.", to=(ANA, BO), cc=())

    assert preview(draft).summary == 'Hello; to ana@example.com, bo@example.com; body: "Hi Ana."'


def test_lists_cc_recipients_when_there_are_any() -> None:
    draft = NewMailDraft(subject="Hello", body="Hi.", to=(ANA,), cc=(BO,))

    assert "cc bo@example.com" in preview(draft).summary


def test_omits_the_cc_clause_when_there_are_none() -> None:
    draft = NewMailDraft(subject="Hello", body="Hi.", to=(ANA,))

    assert "cc" not in preview(draft).summary


def test_says_so_explicitly_when_the_body_is_empty() -> None:
    draft = NewMailDraft(subject="Hello", body="", to=(ANA,))

    assert preview(draft).summary.endswith("body: (none)")


def test_truncates_a_long_body_with_an_explicit_character_count() -> None:
    body = "a" * (MAX_BODY_IN_SUMMARY + 1)
    draft = NewMailDraft(subject="Hello", body=body, to=(ANA,))

    summary = preview(draft).summary

    assert f"{len(body)} characters total" in summary
    assert body not in summary
