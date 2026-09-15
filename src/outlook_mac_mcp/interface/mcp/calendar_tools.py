from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.application.list_upcoming_events import (
    DEFAULT_UPCOMING_DAYS,
    ListUpcomingEvents,
)
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.event_page_view import EventPageView
from outlook_mac_mcp.interface.mcp.list_upcoming_events_input import (
    Days,
    ListUpcomingEventsInput,
)
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.totals_guidance import totals_guidance
from outlook_mac_mcp.interface.mcp.use_cases import UseCases

LIST_TODAYS_EVENTS_TOOL = "list_todays_events"
LIST_TODAYS_EVENTS_DESCRIPTION = (
    "List today's calendar events, earliest first. Today runs from midnight to midnight "
    "in the server's configured time zone, and every time carries its UTC offset. "
    "Includes all-day events and events that started yesterday or end tomorrow. "
    "Recurring series appear as today's occurrence. Returns metadata only. "
    + totals_guidance("this tool takes no arguments, so tell the user the day holds more")
)
LIST_UPCOMING_EVENTS_TOOL = "list_upcoming_events"
LIST_UPCOMING_EVENTS_DESCRIPTION = (
    "List calendar events from now until `days` days ahead, earliest first, with every "
    "time carrying its UTC offset. An event already in progress is included; one that "
    "ended earlier today is not. Recurring series appear as their occurrences. "
    "Returns metadata only. " + totals_guidance("fewer days")
)


def register_calendar_tools(server: MCPServer, use_cases: UseCases) -> None:
    _register_list_todays_events(server, use_cases.list_todays_events)
    _register_list_upcoming_events(server, use_cases.list_upcoming_events)


def _register_list_todays_events(server: MCPServer, use_case: ListTodaysEvents) -> None:
    @server.tool(name=LIST_TODAYS_EVENTS_TOOL, description=LIST_TODAYS_EVENTS_DESCRIPTION)
    async def list_todays_events() -> EventPageView:
        try:
            return _translate_today(use_case)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_today(use_case: ListTodaysEvents) -> EventPageView:
    with observed_tool_call(LIST_TODAYS_EVENTS_TOOL) as outcome:
        page = use_case.execute()
        outcome.item_count = len(page.items)
        return EventPageView.from_page(page)


def _register_list_upcoming_events(server: MCPServer, use_case: ListUpcomingEvents) -> None:
    @server.tool(name=LIST_UPCOMING_EVENTS_TOOL, description=LIST_UPCOMING_EVENTS_DESCRIPTION)
    async def list_upcoming_events(days: Days = DEFAULT_UPCOMING_DAYS) -> EventPageView:
        try:
            return _translate_upcoming(use_case, ListUpcomingEventsInput(days=days))
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_upcoming(
    use_case: ListUpcomingEvents, model: ListUpcomingEventsInput
) -> EventPageView:
    with observed_tool_call(LIST_UPCOMING_EVENTS_TOOL) as outcome:
        page = use_case.execute(model.to_request())
        outcome.item_count = len(page.items)
        return EventPageView.from_page(page)
