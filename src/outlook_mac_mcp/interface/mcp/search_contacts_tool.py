from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT
from outlook_mac_mcp.application.search_contacts import SearchContacts
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.contact_page_view import ContactPageView
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.search_contacts_input import (
    Limit,
    SearchContactsInput,
    Term,
)
from outlook_mac_mcp.interface.mcp.totals_guidance import totals_guidance

SEARCH_CONTACTS_TOOL = "search_contacts"
SEARCH_CONTACTS_DESCRIPTION = (
    "Search the signed-in user's contacts by a term matched two different ways: as a "
    "PREFIX against the contact's display name, and as an EXACT match against any of "
    "the contact's email addresses. Microsoft Graph does not support a prefix match on "
    "email address, only an exact one, so a term that is not a complete address will "
    "match by name only, never by a partial address. "
    + totals_guidance("a longer or more exact term")
)


def register_search_contacts_tool(server: MCPServer, use_case: SearchContacts) -> None:
    @server.tool(name=SEARCH_CONTACTS_TOOL, description=SEARCH_CONTACTS_DESCRIPTION)
    async def search_contacts(term: Term, limit: Limit = DEFAULT_LIMIT) -> ContactPageView:
        try:
            return _translate(use_case, SearchContactsInput(term=term, limit=limit))
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate(use_case: SearchContacts, model: SearchContactsInput) -> ContactPageView:
    with observed_tool_call(SEARCH_CONTACTS_TOOL) as outcome:
        page = use_case.execute(model.to_request())
        outcome.item_count = len(page.items)
        return ContactPageView.from_page(page)
