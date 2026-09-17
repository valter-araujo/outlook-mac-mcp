from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.resolve_folders import ResolveFolders
from outlook_mac_mcp.application.top_senders import (
    DEFAULT_TOP_SENDERS,
    SCAN_CEILING,
    TopSenders,
)
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.interface.mcp.email_filters_input import (
    FilterFolder,
    IsRead,
    ReceivedAfter,
    ReceivedBefore,
)
from outlook_mac_mcp.interface.mcp.folder_scope_guidance import folder_scope_guidance
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.sender_ranking_view import SenderRankingView
from outlook_mac_mcp.interface.mcp.top_senders_input import SenderLimit, TopSendersInput

TOP_SENDERS_TOOL = "top_senders"
TOP_SENDERS_DESCRIPTION = (
    "Rank the senders of a folder by how many emails each sent, most first, over the "
    "emails matching the optional filters (read state, received-time range). The ranking "
    f"is built by scanning the matching emails, at most {SCAN_CEILING:,} per call -- when "
    "folder is all, that ceiling is shared across all five folders combined, not "
    f"{SCAN_CEILING:,} each. The output carries senders (address, name, count), scanned "
    "(emails counted), total (the exact number of emails matching the filters) and "
    "coverage_is_complete. "
    "When coverage_is_complete is false, the ranking covers ONLY the scanned emails, not "
    "the whole folder: tell the user it was computed over scanned of total emails, do not "
    "present it as the folder's ranking, and suggest narrowing with received_after (and "
    "received_before) so that every matching email fits within the scan. "
    "Counts are per scanned email; use count_emails with a sender filter for an exact "
    "count of one sender across the whole folder. " + folder_scope_guidance()
)


def register_top_senders_tool(
    server: MCPServer, use_case: TopSenders, resolve_folders: ResolveFolders
) -> None:
    @server.tool(name=TOP_SENDERS_TOOL, description=TOP_SENDERS_DESCRIPTION)
    async def top_senders(
        folder: FilterFolder = FolderSelection.INBOX,
        is_read: IsRead = None,
        received_after: ReceivedAfter = None,
        received_before: ReceivedBefore = None,
        limit: SenderLimit = DEFAULT_TOP_SENDERS,
    ) -> SenderRankingView:
        model = TopSendersInput(
            folder=folder,
            is_read=is_read,
            received_after=received_after,
            received_before=received_before,
            limit=limit,
        )
        try:
            return _translate(use_case, resolve_folders, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate(
    use_case: TopSenders, resolve_folders: ResolveFolders, model: TopSendersInput
) -> SenderRankingView:
    with observed_tool_call(TOP_SENDERS_TOOL) as outcome:
        resolved = resolve_folders.execute(model.folder)
        ranking = use_case.execute(model.to_request(resolved.folder_ids))
        outcome.item_count = len(ranking.senders)
        return SenderRankingView.from_ranking(ranking, folder=resolved.echo)
