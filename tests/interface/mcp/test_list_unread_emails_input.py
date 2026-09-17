import pytest
from pydantic import ValidationError

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmailsRequest
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.interface.mcp.list_unread_emails_input import ListUnreadEmailsInput

# Folder resolution itself (well-known name, all, custom path, unknown path) is
# ResolveFolders' behavior, tested in tests/application/test_resolve_folders.py; this
# model only carries the raw folder string and passes already-resolved ids through.


def test_defaults_to_the_inbox_argument_and_the_use_case_default_limit() -> None:
    model = ListUnreadEmailsInput()

    assert model.folder == FolderSelection.INBOX
    assert model.to_request(folder_ids=(FolderName.INBOX,)) == ListUnreadEmailsRequest(
        folders=(FolderName.INBOX,), limit=DEFAULT_LIMIT
    )


def test_accepts_an_arbitrary_string_as_the_folder_argument() -> None:
    model = ListUnreadEmailsInput.model_validate({"folder": "Entrevistas/Work/AWS"})

    assert model.folder == "Entrevistas/Work/AWS"


def test_passes_resolved_folder_ids_straight_through() -> None:
    model = ListUnreadEmailsInput(limit=5)

    request = model.to_request(folder_ids=("some-graph-id",))

    assert request.folders == ("some-graph-id",)
    assert request.limit == 5


@pytest.mark.parametrize("limit", [MIN_LIMIT, MAX_LIMIT])
def test_accepts_a_limit_at_the_edge_of_the_allowed_range(limit: int) -> None:
    model = ListUnreadEmailsInput(limit=limit)

    assert model.to_request(folder_ids=(FolderName.INBOX,)).limit == limit


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_allowed_range(limit: int) -> None:
    with pytest.raises(ValidationError):
        ListUnreadEmailsInput(limit=limit)


def test_rejects_an_empty_folder_argument() -> None:
    with pytest.raises(ValidationError):
        ListUnreadEmailsInput.model_validate({"folder": ""})


def test_rejects_an_unexpected_argument() -> None:
    with pytest.raises(ValidationError):
        ListUnreadEmailsInput.model_validate({"unexpected": 1})
