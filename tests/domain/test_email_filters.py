from datetime import UTC, datetime, timedelta

import pytest

from outlook_mac_mcp.domain.email_filters import MAX_SENDER_ADDRESS_LENGTH, EmailFilters
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName

A_MOMENT = datetime(2026, 9, 1, tzinfo=UTC)


def test_defaults_to_the_inbox_with_nothing_else_restricted() -> None:
    filters = EmailFilters()

    assert filters.folders == (FolderName.INBOX,)
    assert filters.is_empty()


def test_rejects_an_empty_folders_tuple() -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(folders=())


def test_rejects_a_repeated_folder() -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(folders=(FolderName.INBOX, FolderName.INBOX))


def test_accepts_every_well_known_folder_at_once() -> None:
    filters = EmailFilters(folders=tuple(FolderName))

    assert set(filters.folders) == set(FolderName)


def test_is_not_empty_once_any_restriction_is_set() -> None:
    assert not EmailFilters(is_read=False).is_empty()
    assert not EmailFilters(has_attachments=False).is_empty()


def test_accepts_a_sender_address() -> None:
    assert EmailFilters(sender="ana@example.com").sender == "ana@example.com"


@pytest.mark.parametrize(
    "sender",
    ["not-an-address", "two@@example.com", "spaced name@example.com", "o'neil@example.com", ""],
)
def test_rejects_a_sender_that_is_not_a_plain_address(sender: str) -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(sender=sender)


def test_rejects_an_overlong_sender() -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(sender="a" * MAX_SENDER_ADDRESS_LENGTH + "@example.com")


def test_accepts_an_ordered_date_range() -> None:
    filters = EmailFilters(received_after=A_MOMENT, received_before=A_MOMENT + timedelta(days=1))

    assert filters.received_after == A_MOMENT


def test_rejects_a_date_range_that_is_empty_or_inverted() -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(received_after=A_MOMENT, received_before=A_MOMENT)
    with pytest.raises(InvalidRequestError):
        EmailFilters(received_after=A_MOMENT + timedelta(days=1), received_before=A_MOMENT)


def test_rejects_a_naive_lower_bound() -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(received_after=datetime(2026, 9, 1))


def test_rejects_a_naive_upper_bound() -> None:
    with pytest.raises(InvalidRequestError):
        EmailFilters(received_before=datetime(2026, 9, 1))
