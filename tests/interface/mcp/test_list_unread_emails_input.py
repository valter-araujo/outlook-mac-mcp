import pytest
from pydantic import ValidationError

from outlook_mac_mcp.application.list_unread_emails import (
    DEFAULT_LIMIT,
    MAX_LIMIT,
    MIN_LIMIT,
    ListUnreadEmailsRequest,
)
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.interface.mcp.list_unread_emails_input import ListUnreadEmailsInput


def test_defaults_to_the_inbox_and_the_use_case_default_limit() -> None:
    request = ListUnreadEmailsInput().to_request()

    assert request == ListUnreadEmailsRequest(folder=FolderName.INBOX, limit=DEFAULT_LIMIT)


def test_translates_the_requested_folder_and_limit() -> None:
    model = ListUnreadEmailsInput(folder=FolderName.ARCHIVE, limit=5)

    assert model.to_request() == ListUnreadEmailsRequest(folder=FolderName.ARCHIVE, limit=5)


def test_accepts_the_folder_as_its_graph_name() -> None:
    model = ListUnreadEmailsInput.model_validate({"folder": "junkemail"})

    assert model.to_request().folder is FolderName.JUNK


@pytest.mark.parametrize("limit", [MIN_LIMIT, MAX_LIMIT])
def test_accepts_a_limit_at_the_edge_of_the_allowed_range(limit: int) -> None:
    assert ListUnreadEmailsInput(limit=limit).to_request().limit == limit


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_allowed_range(limit: int) -> None:
    with pytest.raises(ValidationError):
        ListUnreadEmailsInput(limit=limit)


def test_rejects_an_unknown_folder() -> None:
    with pytest.raises(ValidationError):
        ListUnreadEmailsInput.model_validate({"folder": "nowhere"})


def test_rejects_an_unexpected_argument() -> None:
    with pytest.raises(ValidationError):
        ListUnreadEmailsInput.model_validate({"unexpected": 1})
