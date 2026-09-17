from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.preview_event import PreviewEvent
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.calendar_deletion_tools import register_calendar_deletion_tools
from outlook_mac_mcp.interface.mcp.calendar_update_tools import register_calendar_update_tools
from outlook_mac_mcp.interface.mcp.calendar_write_use_cases import CalendarWriteUseCases
from outlook_mac_mcp.interface.mcp.event_detail_view import EventDetailView
from outlook_mac_mcp.interface.mcp.event_draft_view import EventDraftView
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.preview_event_input import (
    Attendees,
    Body,
    End,
    EventSensitivity,
    EventShowAs,
    IsAllDay,
    Location,
    PreviewEventInput,
    ReminderMinutesBeforeStart,
    Start,
    Subject,
)
from outlook_mac_mcp.interface.mcp.write_confirmation_rule import write_confirmation_rule

PREVIEW_EVENT_TOOL = "preview_event"
CREATE_EVENT_TOOL = "create_event"
CONFIRMATION_RULE = write_confirmation_rule(PREVIEW_EVENT_TOOL)
PREVIEW_EVENT_DESCRIPTION = (
    "Prepare a calendar event without creating it. Validates the details and returns a "
    "token and a one-line summary. Nothing is written to the calendar by this tool. "
    "Show the summary to the user and, only once they agree, pass the token to "
    "create_event. " + CONFIRMATION_RULE
)
CREATE_EVENT_DESCRIPTION = (
    "Create the event previewed under `token`. This performs a REAL CHANGE to the user's "
    "calendar and cannot be undone by this server. The token must come from preview_event "
    "in this same session; each token works exactly once, and an unknown or already used "
    "token is refused. Call this only after the user has confirmed the preview's summary. "
    + CONFIRMATION_RULE
)


def register_calendar_write_tools(server: MCPServer, use_cases: CalendarWriteUseCases) -> None:
    _register_preview_event(server, use_cases.preview_event)
    _register_create_event(server, use_cases.create_event)
    register_calendar_update_tools(server, use_cases)
    register_calendar_deletion_tools(server, use_cases)


def _register_preview_event(server: MCPServer, use_case: PreviewEvent) -> None:
    @server.tool(name=PREVIEW_EVENT_TOOL, description=PREVIEW_EVENT_DESCRIPTION)
    async def preview_event(
        subject: Subject,
        start: Start,
        end: End,
        is_all_day: IsAllDay = False,
        location: Location = "",
        body: Body = "",
        attendees: Attendees = (),
        reminder_minutes_before_start: ReminderMinutesBeforeStart | None = None,
        sensitivity: EventSensitivity | None = None,
        show_as: EventShowAs | None = None,
    ) -> EventDraftView:
        model = PreviewEventInput(
            subject=subject,
            start=start,
            end=end,
            is_all_day=is_all_day,
            location=location,
            body=body,
            attendees=attendees,
            reminder_minutes_before_start=reminder_minutes_before_start,
            sensitivity=sensitivity,
            show_as=show_as,
        )
        try:
            return _translate_preview(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_preview(use_case: PreviewEvent, model: PreviewEventInput) -> EventDraftView:
    with observed_tool_call(PREVIEW_EVENT_TOOL) as outcome:
        draft = use_case.execute(model.to_new_event())
        outcome.item_count = 1
        return EventDraftView.from_draft(draft)


def _register_create_event(server: MCPServer, use_case: CreateEvent) -> None:
    @server.tool(name=CREATE_EVENT_TOOL, description=CREATE_EVENT_DESCRIPTION)
    async def create_event(token: str) -> EventDetailView:
        try:
            return _translate_create(use_case, token)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_create(use_case: CreateEvent, token: str) -> EventDetailView:
    with observed_tool_call(CREATE_EVENT_TOOL) as outcome:
        event = use_case.execute(token)
        outcome.item_count = 1
        return EventDetailView.from_event(event)
