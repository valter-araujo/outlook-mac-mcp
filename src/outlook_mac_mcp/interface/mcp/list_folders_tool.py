from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.list_folders import ListFolders
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.mail_folder_view import MailFolderView
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call

LIST_FOLDERS_TOOL = "list_folders"
LIST_FOLDERS_DESCRIPTION = (
    "List the mailbox's well-known folders (inbox, archive, junk email, sent items, "
    "drafts) with their display name and their live unread and total item counts. "
    "Takes no arguments and always returns the complete set: it is not a search or a "
    "filtered listing. Arbitrary user-created folders are not included; v1 reads only "
    "the well-known set."
)


def register_list_folders_tool(server: MCPServer, use_case: ListFolders) -> None:
    @server.tool(name=LIST_FOLDERS_TOOL, description=LIST_FOLDERS_DESCRIPTION)
    async def list_folders() -> list[MailFolderView]:
        try:
            return _translate(use_case)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate(use_case: ListFolders) -> list[MailFolderView]:
    with observed_tool_call(LIST_FOLDERS_TOOL) as outcome:
        folders = use_case.execute()
        outcome.item_count = len(folders)
        return [MailFolderView.from_folder(folder) for folder in folders]
