import json
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.list_folders_tool import LIST_FOLDERS_TOOL
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_mail_folder_repository import InMemoryMailFolderRepository
from tests.fakes.use_case_bundles import folders_use_cases

pytestmark = pytest.mark.anyio


def make_folder(name: FolderName, *, unread: int = 0, total: int = 0) -> MailFolder:
    return MailFolder(
        well_known_name=name,
        display_name=name.value.title(),
        unread_count=unread,
        total_count=total,
    )


def make_custom_folder(
    path: str, *, folder_id: str = "id", unread: int = 0, total: int = 0
) -> CustomMailFolder:
    return CustomMailFolder(
        folder_id=folder_id,
        display_name=path.rsplit("/", 1)[-1],
        path=path,
        unread_count=unread,
        total_count=total,
    )


class FailingMailFolderRepository:
    def __init__(self, error: Exception) -> None:
        self._error = error

    def list_all(self) -> tuple[MailFolder, ...]:
        raise self._error

    def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
        raise self._error


def server_with(*folders: MailFolder, custom: CustomFolderScan | None = None) -> MCPServer:
    repository = InMemoryMailFolderRepository()
    for folder in folders:
        repository.add(folder)
    if custom is not None:
        repository.set_custom(custom)
    return build_server(folders_use_cases(repository))


async def call(server: MCPServer, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(LIST_FOLDERS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def the_tool(server: MCPServer) -> Any:
    return next(tool for tool in await server.list_tools() if tool.name == LIST_FOLDERS_TOOL)


async def test_takes_no_arguments() -> None:
    tool = await the_tool(server_with())

    assert tool.input_schema.get("properties", {}) == {}


async def test_returns_each_well_known_folder_with_its_name_and_counts() -> None:
    server = server_with(
        make_folder(FolderName.INBOX, unread=70, total=71), make_folder(FolderName.ARCHIVE)
    )

    payload = await call(server, {})

    assert payload["well_known"][0]["well_known_name"] == "inbox"
    assert payload["well_known"][0]["display_name"] == "Inbox"
    assert payload["well_known"][0]["unread_count"] == 70
    assert payload["well_known"][0]["total_count"] == 71


async def test_returns_empty_lists_when_the_repository_holds_nothing() -> None:
    payload = await call(server_with(), {})

    assert payload["well_known"] == []
    assert payload["custom"] == []
    assert payload["depth_limit_reached"] is False
    assert payload["folder_limit_reached"] is False


async def test_returns_each_custom_folder_with_its_path_and_counts() -> None:
    custom = CustomFolderScan(
        folders=(make_custom_folder("Candidaturas/2026", folder_id="f1", unread=3, total=9),),
        depth_limit_reached=False,
        folder_limit_reached=False,
    )
    server = server_with(custom=custom)

    payload = await call(server, {})

    assert payload["custom"] == [
        {
            "folder_id": "f1",
            "display_name": "2026",
            "path": "Candidaturas/2026",
            "unread_count": 3,
            "total_count": 9,
        }
    ]


async def test_keeps_two_custom_folders_with_the_same_name_distinct_by_path() -> None:
    custom = CustomFolderScan(
        folders=(
            make_custom_folder("Candidaturas/2026", folder_id="f1"),
            make_custom_folder("Projetos/2026", folder_id="f2"),
        ),
        depth_limit_reached=False,
        folder_limit_reached=False,
    )
    server = server_with(custom=custom)

    payload = await call(server, {})

    paths = {folder["path"] for folder in payload["custom"]}
    assert paths == {"Candidaturas/2026", "Projetos/2026"}


async def test_reports_when_the_depth_or_folder_limit_was_reached() -> None:
    custom = CustomFolderScan(folders=(), depth_limit_reached=True, folder_limit_reached=True)
    server = server_with(custom=custom)

    payload = await call(server, {})

    assert payload["depth_limit_reached"] is True
    assert payload["folder_limit_reached"] is True


async def test_translates_a_project_error_into_a_tool_error() -> None:
    server = build_server(
        folders_use_cases(FailingMailFolderRepository(NotAuthenticatedError("sign in first")))
    )

    with pytest.raises(ToolError, match="sign in first"):
        await call(server, {})


async def test_describes_both_well_known_and_custom_sections() -> None:
    description = (await the_tool(server_with())).description or ""

    assert "well_known" in description
    assert "custom" in description
    assert "full path" in description


async def test_logs_the_call_with_the_combined_item_count(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()
    custom = CustomFolderScan(
        folders=(make_custom_folder("Candidaturas"),),
        depth_limit_reached=False,
        folder_limit_reached=False,
    )
    server = server_with(
        make_folder(FolderName.INBOX), make_folder(FolderName.ARCHIVE), custom=custom
    )

    await call(server, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == LIST_FOLDERS_TOOL
    assert record["item_count"] == 3
