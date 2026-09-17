from outlook_mac_mcp.interface.mcp.calendar_tools import (
    LIST_TODAYS_EVENTS_DESCRIPTION,
    LIST_UPCOMING_EVENTS_DESCRIPTION,
)
from outlook_mac_mcp.interface.mcp.list_largest_emails_tool import LIST_LARGEST_EMAILS_DESCRIPTION
from outlook_mac_mcp.interface.mcp.mail_listing_tools import (
    COUNT_EMAILS_DESCRIPTION,
    LIST_EMAILS_DESCRIPTION,
)
from outlook_mac_mcp.interface.mcp.search_contacts_tool import SEARCH_CONTACTS_DESCRIPTION
from outlook_mac_mcp.interface.mcp.server import (
    GET_EMAIL_DESCRIPTION,
    LIST_UNREAD_EMAILS_DESCRIPTION,
    SEARCH_EMAILS_DESCRIPTION,
)
from outlook_mac_mcp.interface.mcp.top_senders_tool import TOP_SENDERS_DESCRIPTION
from outlook_mac_mcp.interface.mcp.untrusted_content_warning import (
    light_untrusted_content_note,
    untrusted_content_warning,
)

FULL_WARNING_DESCRIPTIONS = (
    GET_EMAIL_DESCRIPTION,
    LIST_UNREAD_EMAILS_DESCRIPTION,
    SEARCH_EMAILS_DESCRIPTION,
    LIST_EMAILS_DESCRIPTION,
    TOP_SENDERS_DESCRIPTION,
    LIST_LARGEST_EMAILS_DESCRIPTION,
)
LIGHT_NOTE_DESCRIPTIONS = (
    LIST_TODAYS_EVENTS_DESCRIPTION,
    LIST_UPCOMING_EVENTS_DESCRIPTION,
    SEARCH_CONTACTS_DESCRIPTION,
)


def test_the_full_warning_appears_verbatim_in_every_tool_returning_email_content() -> None:
    warning = untrusted_content_warning()

    for description in FULL_WARNING_DESCRIPTIONS:
        assert warning in description


def test_the_light_note_appears_verbatim_in_every_tool_returning_lower_risk_names() -> None:
    note = light_untrusted_content_note()

    for description in LIGHT_NOTE_DESCRIPTIONS:
        assert note in description


def test_count_emails_carries_neither_warning_since_it_returns_no_content() -> None:
    warning = untrusted_content_warning()
    note = light_untrusted_content_note()

    assert warning not in COUNT_EMAILS_DESCRIPTION
    assert note not in COUNT_EMAILS_DESCRIPTION
