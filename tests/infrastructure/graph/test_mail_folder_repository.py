from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.mail_folder_repository import GraphMailFolderRepository

INBOX_URL = f"{GRAPH_BASE_URL}/me/mailFolders/inbox"
ARCHIVE_URL = f"{GRAPH_BASE_URL}/me/mailFolders/archive"
JUNK_URL = f"{GRAPH_BASE_URL}/me/mailFolders/junkemail"
SENT_URL = f"{GRAPH_BASE_URL}/me/mailFolders/sentitems"
DRAFTS_URL = f"{GRAPH_BASE_URL}/me/mailFolders/drafts"


def folder_payload(display_name: str, *, unread: int = 0, total: int = 0) -> dict[str, Any]:
    return {"displayName": display_name, "unreadItemCount": unread, "totalItemCount": total}


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def repository() -> GraphMailFolderRepository:
    return GraphMailFolderRepository(GraphClient(FakeTokenProvider()))


def mock_every_well_known_folder(**overrides: dict[str, Any]) -> None:
    defaults = {
        INBOX_URL: folder_payload("Inbox"),
        ARCHIVE_URL: folder_payload("Archive"),
        JUNK_URL: folder_payload("Junk Email"),
        SENT_URL: folder_payload("Sent Items"),
        DRAFTS_URL: folder_payload("Drafts"),
    }
    defaults.update(overrides)
    for url, payload in defaults.items():
        respx.get(url).mock(return_value=httpx.Response(200, json=payload))


def test_satisfies_the_mail_folder_repository_port(repository: GraphMailFolderRepository) -> None:
    port: MailFolderRepository = repository

    assert port is repository


@respx.mock
def test_fetches_every_well_known_folder_by_its_own_path(
    repository: GraphMailFolderRepository,
) -> None:
    mock_every_well_known_folder()

    folders = repository.list_all()

    assert [folder.well_known_name for folder in folders] == [
        FolderName.INBOX,
        FolderName.ARCHIVE,
        FolderName.JUNK,
        FolderName.SENT,
        FolderName.DRAFTS,
    ]


@respx.mock
def test_selects_only_the_fields_the_entity_needs(repository: GraphMailFolderRepository) -> None:
    route = respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json=folder_payload("Inbox"))
    )
    mock_every_well_known_folder()

    repository.list_all()

    selected = set(dict(route.calls.last.request.url.params)["$select"].split(","))
    assert selected == {"displayName", "unreadItemCount", "totalItemCount"}


@respx.mock
def test_maps_the_display_name_and_counts_from_each_folder(
    repository: GraphMailFolderRepository,
) -> None:
    mock_every_well_known_folder(**{INBOX_URL: folder_payload("Inbox", unread=70, total=71)})

    folders = repository.list_all()

    inbox = next(folder for folder in folders if folder.well_known_name is FolderName.INBOX)
    assert inbox.display_name == "Inbox"
    assert inbox.unread_count == 70
    assert inbox.total_count == 71


@respx.mock
def test_uses_the_mailboxs_own_localized_display_name(
    repository: GraphMailFolderRepository,
) -> None:
    mock_every_well_known_folder(**{INBOX_URL: folder_payload("Caixa de Entrada")})

    folders = repository.list_all()

    inbox = next(folder for folder in folders if folder.well_known_name is FolderName.INBOX)
    assert inbox.display_name == "Caixa de Entrada"


@respx.mock
def test_raises_when_a_count_is_missing(repository: GraphMailFolderRepository) -> None:
    payload: dict[str, Any] = {"displayName": "Inbox"}
    mock_every_well_known_folder(**{INBOX_URL: payload})

    with pytest.raises(GraphResponseError):
        repository.list_all()


@respx.mock
def test_raises_when_a_count_is_negative(repository: GraphMailFolderRepository) -> None:
    mock_every_well_known_folder(**{INBOX_URL: folder_payload("Inbox", unread=-1)})

    with pytest.raises(GraphResponseError):
        repository.list_all()
