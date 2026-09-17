from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.count_emails import CountEmails
from outlook_mac_mcp.application.limits import DEFAULT_LIMIT
from outlook_mac_mcp.application.list_emails import ListEmails
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.domain.sort_order import SortOrder
from outlook_mac_mcp.interface.mcp.email_count_view import EmailCountView
from outlook_mac_mcp.interface.mcp.email_filters_input import (
    EmailFiltersInput,
    FilterFolder,
    HasAttachments,
    IsRead,
    ReceivedAfter,
    ReceivedBefore,
    Sender,
)
from outlook_mac_mcp.interface.mcp.email_page_view import EmailPageView
from outlook_mac_mcp.interface.mcp.folder_scope_guidance import folder_scope_guidance
from outlook_mac_mcp.interface.mcp.list_emails_input import ListEmailsInput, ListLimit, Sort
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.totals_guidance import totals_guidance
from outlook_mac_mcp.interface.mcp.use_cases import UseCases

LIST_EMAILS_TOOL = "list_emails"
LIST_EMAILS_DESCRIPTION = (
    "List emails in a folder by date, filtered any way you like: read state, exact sender "
    "address, a received-time range (after inclusive, before exclusive), attachments. "
    "Every filter is optional and they combine. sort=newest gives the most recent first; "
    "sort=oldest gives the earliest first, which is how to find a folder's oldest email "
    "or its date range. The total is exact and counts every email matching the filters, "
    "so this tool, unlike search_emails, can answer how many and how far back. "
    "Returns metadata and a short preview; use get_email for a full body. "
    + totals_guidance("a sender, a tighter date range, or a read or attachment state")
    + " "
    + folder_scope_guidance()
)
COUNT_EMAILS_TOOL = "count_emails"
COUNT_EMAILS_DESCRIPTION = (
    "Count the emails in a folder matching the same filters as list_emails, returning only "
    "the exact total and no emails. Use it to answer how many, or to size a question before "
    "listing. It cannot tell which emails match or when they arrived; use list_emails for "
    "that. " + folder_scope_guidance()
)


def register_mail_listing_tools(server: MCPServer, use_cases: UseCases) -> None:
    _register_list_emails(server, use_cases.list_emails)
    _register_count_emails(server, use_cases.count_emails)


def _register_list_emails(server: MCPServer, use_case: ListEmails) -> None:
    @server.tool(name=LIST_EMAILS_TOOL, description=LIST_EMAILS_DESCRIPTION)
    async def list_emails(
        folder: FilterFolder = FolderSelection.INBOX,
        is_read: IsRead = None,
        sender: Sender = None,
        received_after: ReceivedAfter = None,
        received_before: ReceivedBefore = None,
        has_attachments: HasAttachments = None,
        sort: Sort = SortOrder.NEWEST,
        limit: ListLimit = DEFAULT_LIMIT,
    ) -> EmailPageView:
        model = ListEmailsInput(
            folder=folder,
            is_read=is_read,
            sender=sender,
            received_after=received_after,
            received_before=received_before,
            has_attachments=has_attachments,
            sort=sort,
            limit=limit,
        )
        try:
            return _translate_list(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_list(use_case: ListEmails, model: ListEmailsInput) -> EmailPageView:
    with observed_tool_call(LIST_EMAILS_TOOL) as outcome:
        page = use_case.execute(model.to_request())
        outcome.item_count = len(page.items)
        return EmailPageView.from_page(page, folder=model.folder.describe_folders())


def _register_count_emails(server: MCPServer, use_case: CountEmails) -> None:
    @server.tool(name=COUNT_EMAILS_TOOL, description=COUNT_EMAILS_DESCRIPTION)
    async def count_emails(
        folder: FilterFolder = FolderSelection.INBOX,
        is_read: IsRead = None,
        sender: Sender = None,
        received_after: ReceivedAfter = None,
        received_before: ReceivedBefore = None,
        has_attachments: HasAttachments = None,
    ) -> EmailCountView:
        model = EmailFiltersInput(
            folder=folder,
            is_read=is_read,
            sender=sender,
            received_after=received_after,
            received_before=received_before,
            has_attachments=has_attachments,
        )
        try:
            return _translate_count(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_count(use_case: CountEmails, model: EmailFiltersInput) -> EmailCountView:
    """The logged item count is the total itself: it is the one number the call produced."""
    with observed_tool_call(COUNT_EMAILS_TOOL) as outcome:
        total = use_case.execute(model.to_filters())
        outcome.item_count = total
        return EmailCountView(total=total, folder=model.folder.describe_folders())
