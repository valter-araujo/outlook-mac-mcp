from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.list_largest_emails import (
    DEFAULT_LARGEST_EMAILS,
    SCAN_CEILING,
    ListLargestEmails,
)
from outlook_mac_mcp.application.resolve_folders import ResolveFolders
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.interface.mcp.email_filters_input import FilterFolder, IsRead, ReceivedAfter
from outlook_mac_mcp.interface.mcp.email_size_ranking_view import EmailSizeRankingView
from outlook_mac_mcp.interface.mcp.folder_scope_guidance import folder_scope_guidance
from outlook_mac_mcp.interface.mcp.list_largest_emails_input import (
    LargestEmailsLimit,
    ListLargestEmailsInput,
)
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.untrusted_content_warning import untrusted_content_warning

LIST_LARGEST_EMAILS_TOOL = "list_largest_emails"
LIST_LARGEST_EMAILS_DESCRIPTION = (
    "List the largest emails in a folder by byte size, largest first, over the emails "
    "matching the optional filters (read state, received-time range). Size comes from a "
    "Graph extended property (the message's stored byte size), not an estimate; a small "
    "number of emails may carry no such property and are excluded from the ranking, "
    "counted separately in skipped rather than treated as size zero. Graph has no way to "
    "sort by this property, so the ranking is built by scanning the matching emails, at "
    f"most {SCAN_CEILING:,} per call -- when folder is all, that ceiling is shared across "
    f"all five folders combined, not {SCAN_CEILING:,} each. The output carries items (id, "
    "subject, sender, received_at, size_bytes), scanned (emails checked), skipped (of "
    "those, how many had no determinable size and are not in items), total (the exact "
    "number of emails matching the filters) and coverage_is_complete. "
    "When coverage_is_complete is false, the ranking covers ONLY the scanned emails, not "
    "the whole folder: tell the user it was computed over scanned of total emails, do not "
    "present it as the folder's largest emails, and suggest narrowing with received_after "
    "so that every matching email fits within the scan. "
    + folder_scope_guidance()
    + " "
    + untrusted_content_warning()
)


def register_list_largest_emails_tool(
    server: MCPServer, use_case: ListLargestEmails, resolve_folders: ResolveFolders
) -> None:
    @server.tool(name=LIST_LARGEST_EMAILS_TOOL, description=LIST_LARGEST_EMAILS_DESCRIPTION)
    async def list_largest_emails(
        folder: FilterFolder = FolderSelection.INBOX,
        is_read: IsRead = None,
        received_after: ReceivedAfter = None,
        limit: LargestEmailsLimit = DEFAULT_LARGEST_EMAILS,
    ) -> EmailSizeRankingView:
        model = ListLargestEmailsInput(
            folder=folder, is_read=is_read, received_after=received_after, limit=limit
        )
        try:
            return _translate(use_case, resolve_folders, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate(
    use_case: ListLargestEmails, resolve_folders: ResolveFolders, model: ListLargestEmailsInput
) -> EmailSizeRankingView:
    with observed_tool_call(LIST_LARGEST_EMAILS_TOOL) as outcome:
        resolved = resolve_folders.execute(model.folder)
        ranking = use_case.execute(model.to_request(resolved.folder_ids))
        outcome.item_count = len(ranking.items)
        return EmailSizeRankingView.from_ranking(ranking, folder=resolved.echo)
