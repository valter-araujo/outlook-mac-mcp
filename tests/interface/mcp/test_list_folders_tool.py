import json
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
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


class FailingMailFolderRepository:
    def __init__(self, error: Exception) -> None:
        self._error = error

    def list_all(self) -> tuple[MailFolder, ...]:
        raise self._error

    def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
        raise self._error


def server_with(*folders: MailFolder) -> MCPServer:
    repository = InMemoryMailFolderRepository()
    for folder in folders:
        repository.add(folder)
    return build_server(folders_use_cases(repository))


async def call(server: MCPServer, arguments: dict[str, Any]) -> list[dict[str, Any]]:
    result = await server.call_tool(LIST_FOLDERS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    items = result.structured_content["result"]
    assert isinstance(items, list)
    return items


async def the_tool(server: MCPServer) -> Any:
    return next(tool for tool in await server.list_tools() if tool.name == LIST_FOLDERS_TOOL)


async def test_takes_no_arguments() -> None:
    tool = await the_tool(server_with())

    assert tool.input_schema.get("properties", {}) == {}


async def test_returns_each_folder_with_its_name_and_counts() -> None:
    server = server_with(
        make_folder(FolderName.INBOX, unread=70, total=71), make_folder(FolderName.ARCHIVE)
    )

    folders = await call(server, {})

    assert folders[0]["well_known_name"] == "inbox"
    assert folders[0]["display_name"] == "Inbox"
    assert folders[0]["unread_count"] == 70
    assert folders[0]["total_count"] == 71


async def test_returns_an_empty_list_when_the_repository_holds_nothing() -> None:
    assert await call(server_with(), {}) == []


async def test_translates_a_project_error_into_a_tool_error() -> None:
    server = build_server(
        folders_use_cases(FailingMailFolderRepository(NotAuthenticatedError("sign in first")))
    )

    with pytest.raises(ToolError, match="sign in first"):
        await call(server, {})


async def test_says_it_always_returns_the_complete_well_known_set() -> None:
    description = (await the_tool(server_with())).description or ""

    assert "complete set" in description
    assert "not a search" in description
    assert "user-created folders are not included" in description.lower()


async def test_logs_the_call_with_its_item_count(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()
    server = server_with(make_folder(FolderName.INBOX), make_folder(FolderName.ARCHIVE))

    await call(server, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == LIST_FOLDERS_TOOL
    assert record["item_count"] == 2
