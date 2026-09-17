import json
from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.custom_folder_scan import CUSTOM_FOLDER_SCAN_EVENT
from outlook_mac_mcp.infrastructure.graph.mail_folder_repository import GraphMailFolderRepository
from outlook_mac_mcp.interface.mcp.observability import configure_logging

WELL_KNOWN_URLS = {
    "inbox": f"{GRAPH_BASE_URL}/me/mailFolders/inbox",
    "archive": f"{GRAPH_BASE_URL}/me/mailFolders/archive",
    "junkemail": f"{GRAPH_BASE_URL}/me/mailFolders/junkemail",
    "sentitems": f"{GRAPH_BASE_URL}/me/mailFolders/sentitems",
    "drafts": f"{GRAPH_BASE_URL}/me/mailFolders/drafts",
}
WELL_KNOWN_IDS = {
    "inbox": "wk-inbox",
    "archive": "wk-archive",
    "junkemail": "wk-junk",
    "sentitems": "wk-sent",
    "drafts": "wk-drafts",
}


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def repository() -> GraphMailFolderRepository:
    return GraphMailFolderRepository(GraphClient(FakeTokenProvider()))


def folder_item(
    folder_id: str, display_name: str, *, unread: int = 0, total: int = 0, child_count: int = 0
) -> dict[str, Any]:
    return {
        "id": folder_id,
        "displayName": display_name,
        "unreadItemCount": unread,
        "totalItemCount": total,
        "childFolderCount": child_count,
    }


def page(items: list[dict[str, Any]]) -> dict[str, Any]:
    return {"value": items}


def children_url(folder_id: str) -> str:
    return f"{GRAPH_BASE_URL}/me/mailFolders/{folder_id}/childFolders"


def mock_well_known_summaries() -> None:
    for well_known_name, url in WELL_KNOWN_URLS.items():
        payload = {"id": WELL_KNOWN_IDS[well_known_name], "displayName": well_known_name.title()}
        respx.get(url).mock(return_value=httpx.Response(200, json=payload))


def mock_children(folder_id: str, payload: dict[str, Any]) -> None:
    respx.get(children_url(folder_id)).mock(return_value=httpx.Response(200, json=payload))


def mock_empty_children_for_every_well_known() -> None:
    for folder_id in WELL_KNOWN_IDS.values():
        mock_children(folder_id, page([]))


def test_satisfies_the_mail_folder_repository_port(repository: GraphMailFolderRepository) -> None:
    port: MailFolderRepository = repository

    assert port is repository


@respx.mock
def test_discovers_nested_custom_folders_at_least_three_levels_deep(
    repository: GraphMailFolderRepository,
) -> None:
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page(
            [
                folder_item("wk-inbox", "Inbox"),
                folder_item("c-candidaturas", "Candidaturas", child_count=1),
            ]
        ),
    )
    mock_empty_children_for_every_well_known()
    mock_children("c-candidaturas", page([folder_item("c-2026", "2026", child_count=1)]))
    mock_children("c-2026", page([folder_item("c-2026-enviadas", "Enviadas")]))

    scan = repository.list_custom(max_depth=10, max_folders=200)

    paths = {folder.path for folder in scan.folders}
    assert paths == {"Candidaturas", "Candidaturas/2026", "Candidaturas/2026/Enviadas"}
    assert scan.depth_limit_reached is False
    assert scan.folder_limit_reached is False


@respx.mock
def test_disambiguates_same_named_folders_by_full_path(
    repository: GraphMailFolderRepository,
) -> None:
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page(
            [
                folder_item("c-candidaturas", "Candidaturas", child_count=1),
                folder_item("c-projetos", "Projetos", child_count=1),
            ]
        ),
    )
    mock_empty_children_for_every_well_known()
    mock_children("c-candidaturas", page([folder_item("c-cand-2026", "2026")]))
    mock_children("c-projetos", page([folder_item("c-proj-2026", "2026")]))

    scan = repository.list_custom(max_depth=10, max_folders=200)

    paths = {folder.path for folder in scan.folders}
    assert paths == {"Candidaturas", "Projetos", "Candidaturas/2026", "Projetos/2026"}
    duplicated_name = [folder for folder in scan.folders if folder.display_name == "2026"]
    assert len(duplicated_name) == 2
    assert {folder.folder_id for folder in duplicated_name} == {"c-cand-2026", "c-proj-2026"}


@respx.mock
def test_skips_well_known_folders_found_among_the_roots_children(
    repository: GraphMailFolderRepository,
) -> None:
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page([folder_item("wk-inbox", "Inbox"), folder_item("wk-archive", "Archive")]),
    )
    mock_empty_children_for_every_well_known()

    scan = repository.list_custom(max_depth=10, max_folders=200)

    assert scan.folders == ()


@respx.mock
def test_stops_at_the_depth_limit_and_reports_it(repository: GraphMailFolderRepository) -> None:
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page([folder_item("c-candidaturas", "Candidaturas", child_count=1)]),
    )
    mock_empty_children_for_every_well_known()
    # c-candidaturas's own children are deliberately left unmocked: a depth cap of 1
    # must stop before that call is ever made.

    scan = repository.list_custom(max_depth=1, max_folders=200)

    assert [folder.path for folder in scan.folders] == ["Candidaturas"]
    assert scan.depth_limit_reached is True
    assert scan.folder_limit_reached is False


@respx.mock
def test_stops_at_the_folder_limit_and_reports_it(repository: GraphMailFolderRepository) -> None:
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page(
            [
                folder_item("c-a", "A"),
                folder_item("c-b", "B"),
                folder_item("c-c", "C"),
            ]
        ),
    )
    mock_empty_children_for_every_well_known()

    scan = repository.list_custom(max_depth=10, max_folders=2)

    assert len(scan.folders) == 2
    assert scan.folder_limit_reached is True


@respx.mock
def test_a_folder_id_with_url_reserved_characters_is_percent_encoded(
    repository: GraphMailFolderRepository,
) -> None:
    """A discovered folder's id becomes the next parent_id when the walk recurses into
    it -- proven here with an id that actually has reserved characters, rather than one
    that happens not to need encoding.
    """
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page([folder_item("id/with#reserved?chars", "Reserved", child_count=1)]),
    )
    mock_empty_children_for_every_well_known()
    encoded_url = f"{GRAPH_BASE_URL}/me/mailFolders/id%2Fwith%23reserved%3Fchars/childFolders"
    route = respx.get(encoded_url).mock(return_value=httpx.Response(200, json=page([])))

    repository.list_custom(max_depth=10, max_folders=200)

    assert route.call_count == 1


@respx.mock
def test_logs_pages_and_folders_found_but_never_a_display_name(
    repository: GraphMailFolderRepository, capsys: pytest.CaptureFixture[str]
) -> None:
    configure_logging()
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page([folder_item("c-candidaturas", "Candidaturas", child_count=1)]),
    )
    mock_empty_children_for_every_well_known()
    mock_children("c-candidaturas", page([folder_item("c-2026", "2026", child_count=0)]))

    repository.list_custom(max_depth=10, max_folders=200)

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["event"] == CUSTOM_FOLDER_SCAN_EVENT
    assert record["folders_found"] == 2
    # msgfolderroot + 5 well-known folders + c-candidaturas's own childFolders call.
    assert record["pages"] == 7
    assert record["depth_limit_reached"] is False
    assert record["folder_limit_reached"] is False
    assert isinstance(record["duration_ms"], float)
    assert "Candidaturas" not in captured.err


@respx.mock
def test_logs_when_the_depth_limit_was_reached(
    repository: GraphMailFolderRepository, capsys: pytest.CaptureFixture[str]
) -> None:
    configure_logging()
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page([folder_item("c-candidaturas", "Candidaturas", child_count=1)]),
    )
    mock_empty_children_for_every_well_known()

    repository.list_custom(max_depth=1, max_folders=200)

    record = json.loads(capsys.readouterr().err.splitlines()[-1])
    assert record["depth_limit_reached"] is True
    assert record["folder_limit_reached"] is False


@respx.mock
def test_logs_when_the_folder_limit_was_reached(
    repository: GraphMailFolderRepository, capsys: pytest.CaptureFixture[str]
) -> None:
    configure_logging()
    mock_well_known_summaries()
    mock_children(
        "msgfolderroot",
        page([folder_item("c-a", "A"), folder_item("c-b", "B")]),
    )
    mock_empty_children_for_every_well_known()

    repository.list_custom(max_depth=10, max_folders=1)

    record = json.loads(capsys.readouterr().err.splitlines()[-1])
    assert record["depth_limit_reached"] is False
    assert record["folder_limit_reached"] is True
