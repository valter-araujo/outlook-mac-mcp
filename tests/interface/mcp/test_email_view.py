import json
from datetime import UTC, datetime

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.interface.mcp.email_view import EmailView

AN_EMAIL = Email(
    id="AAMkAGI2",
    subject="Quarterly review",
    sender=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
    received_at=datetime(2026, 9, 14, 12, 30, tzinfo=UTC),
    is_read=False,
    has_attachments=True,
    preview="Attached is the deck",
)


def test_projects_every_field_of_the_entity() -> None:
    view = EmailView.from_email(AN_EMAIL)

    assert view.id == "AAMkAGI2"
    assert view.subject == "Quarterly review"
    assert view.sender_address == "ana@example.com"
    assert view.sender_name == "Ana Lima"
    assert view.received_at == datetime(2026, 9, 14, 12, 30, tzinfo=UTC)
    assert view.is_read is False
    assert view.has_attachments is True
    assert view.preview == "Attached is the deck"


def test_serializes_to_json_with_an_iso_8601_timestamp() -> None:
    payload = json.loads(EmailView.from_email(AN_EMAIL).model_dump_json())

    assert payload["received_at"] == "2026-09-14T12:30:00Z"


def test_keeps_an_empty_sender_empty() -> None:
    email = Email(
        id="1",
        subject="",
        sender=EmailAddress(address=""),
        received_at=datetime(2026, 9, 14, tzinfo=UTC),
        is_read=False,
        has_attachments=False,
        preview="",
    )

    view = EmailView.from_email(email)

    assert view.sender_address == ""
    assert view.sender_name == ""
