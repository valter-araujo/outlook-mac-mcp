from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.list_custom_folders import ListCustomFolders
from outlook_mac_mcp.application.list_folders import ListFolders
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.list_folders_view import ListFoldersView
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call

LIST_FOLDERS_TOOL = "list_folders"
LIST_FOLDERS_DESCRIPTION = (
    "List the mailbox's folders. `well_known` is the five well-known folders (inbox, "
    "archive, junk email, sent items, drafts) with their display name and their live "
    "unread and total item counts -- always the complete set, this part is not a search "
    "or a filtered listing. `custom` is every user-created folder found by walking the "
    "mailbox's folder tree, each with its full path (e.g. 'Candidaturas/2026') since two "
    "folders can share a display name; address a folder by path, not display name alone. "
    "The walk is capped in depth and in how many folders it visits in one call; "
    "depth_limit_reached and folder_limit_reached say whether either cap cut the search "
    "short, in which case custom may not hold every folder in the mailbox. Graph does not "
    "flag a folder as well-known, so a Microsoft-managed folder that is not one of the "
    "five (e.g. Deleted Items) is reported the same as a genuine custom folder. Takes no "
    "arguments."
)


def register_list_folders_tool(
    server: MCPServer, list_folders: ListFolders, list_custom_folders: ListCustomFolders
) -> None:
    @server.tool(name=LIST_FOLDERS_TOOL, description=LIST_FOLDERS_DESCRIPTION)
    async def list_folders_tool() -> ListFoldersView:
        try:
            return _translate(list_folders, list_custom_folders)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate(
    list_folders: ListFolders, list_custom_folders: ListCustomFolders
) -> ListFoldersView:
    with observed_tool_call(LIST_FOLDERS_TOOL) as outcome:
        well_known = list_folders.execute()
        custom = list_custom_folders.execute()
        outcome.item_count = len(well_known) + len(custom.folders)
        return ListFoldersView.from_result(well_known, custom)
