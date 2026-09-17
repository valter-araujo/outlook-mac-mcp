import json
from datetime import UTC, datetime

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.email_page_view import EmailPageView


def make_email(email_id: str) -> Email:
    return Email(
        id=email_id,
        subject="Quarterly review",
        sender=EmailAddress(address="ana@example.com"),
        received_at=datetime(2026, 9, 14, 12, 30, tzinfo=UTC),
        is_read=False,
        has_attachments=False,
        preview="",
    )


def test_projects_the_items_and_counts_them() -> None:
    view = EmailPageView.from_page(
        Page(items=(make_email("a"), make_email("b")), total=9, total_is_exact=True),
        folder="inbox",
    )

    assert [item.id for item in view.items] == ["a", "b"]
    assert view.returned == 2
    assert view.total == 9
    assert view.total_is_exact is True


def test_keeps_an_inexact_total_flagged_as_such() -> None:
    view = EmailPageView.from_page(
        Page(items=(make_email("a"),), total=250, total_is_exact=False), folder="inbox"
    )

    assert view.total == 250
    assert view.total_is_exact is False


def test_projects_an_empty_page_as_zero_of_zero() -> None:
    view = EmailPageView.from_page(Page(items=(), total=0, total_is_exact=True), folder="inbox")

    assert view.items == []
    assert view.returned == 0
    assert view.total == 0


def test_carries_the_folder_it_was_scoped_to() -> None:
    view = EmailPageView.from_page(Page(items=(), total=0, total_is_exact=True), folder="all")

    assert view.folder == "all"


def test_serializes_the_totals_alongside_the_items() -> None:
    page: Page[Email] = Page(items=(make_email("a"),), total=3, total_is_exact=True)

    payload = json.loads(EmailPageView.from_page(page, folder="inbox").model_dump_json())

    assert set(payload) == {"items", "returned", "total", "total_is_exact", "folder"}
    assert payload["items"][0]["id"] == "a"
