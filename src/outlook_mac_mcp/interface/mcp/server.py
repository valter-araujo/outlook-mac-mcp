from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.list_unread_emails import DEFAULT_LIMIT, ListUnreadEmails
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.interface.mcp.email_view import EmailView
from outlook_mac_mcp.interface.mcp.list_unread_emails_input import (
    Folder,
    Limit,
    ListUnreadEmailsInput,
)
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call

SERVER_NAME = "outlook-mac-mcp"
SERVER_VERSION = "0.1.0"
LIST_UNREAD_EMAILS_TOOL = "list_unread_emails"
LIST_UNREAD_EMAILS_DESCRIPTION = (
    "List unread emails from a mailbox folder, newest first. "
    "Returns metadata and a short preview, never the full body."
)


def build_server(use_case: ListUnreadEmails) -> MCPServer:
    server = MCPServer(name=SERVER_NAME, version=SERVER_VERSION)
    _register_list_unread_emails(server, use_case)
    return server


def _register_list_unread_emails(server: MCPServer, use_case: ListUnreadEmails) -> None:
    @server.tool(name=LIST_UNREAD_EMAILS_TOOL, description=LIST_UNREAD_EMAILS_DESCRIPTION)
    async def list_unread_emails(
        folder: Folder = FolderName.INBOX,
        limit: Limit = DEFAULT_LIMIT,
    ) -> list[EmailView]:
        try:
            return _translate(use_case, ListUnreadEmailsInput(folder=folder, limit=limit))
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate(use_case: ListUnreadEmails, model: ListUnreadEmailsInput) -> list[EmailView]:
    """Observe from inside the error translation, so the log names the project error type
    rather than the ToolError it is about to become.
    """
    with observed_tool_call(LIST_UNREAD_EMAILS_TOOL) as outcome:
        emails = use_case.execute(model.to_request())
        outcome.item_count = len(emails)
        return [EmailView.from_email(email) for email in emails]
