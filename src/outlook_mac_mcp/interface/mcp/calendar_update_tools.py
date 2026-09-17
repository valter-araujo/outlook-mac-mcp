from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.preview_event_update import PreviewEventUpdate
from outlook_mac_mcp.application.update_event import UpdateEvent
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.calendar_write_use_cases import CalendarWriteUseCases
from outlook_mac_mcp.interface.mcp.event_detail_view import EventDetailView
from outlook_mac_mcp.interface.mcp.event_update_draft_view import EventUpdateDraftView
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.update_event_input import (
    Attendees,
    Body,
    End,
    EventId,
    EventSensitivity,
    EventShowAs,
    Location,
    ReminderMinutesBeforeStart,
    Start,
    Subject,
    UpdateEventInput,
)
from outlook_mac_mcp.interface.mcp.write_confirmation_rule import write_confirmation_rule

PREVIEW_EVENT_UPDATE_TOOL = "preview_event_update"
UPDATE_EVENT_TOOL = "update_event"
CONFIRMATION_RULE = write_confirmation_rule(PREVIEW_EVENT_UPDATE_TOOL)
PREVIEW_EVENT_UPDATE_DESCRIPTION = (
    "Prepare a change to an existing event without applying it. Only the fields you "
    "pass are changed; every field you leave out is left exactly as it is. Fetches the "
    "current event and returns a token and a summary showing the diff for each changed "
    "field as old value -> new value, never just the resulting state. Nothing is "
    "written to the calendar by this tool. Show the summary to the user and, only once "
    "they agree, pass the token to update_event. " + CONFIRMATION_RULE
)
UPDATE_EVENT_DESCRIPTION = (
    "Apply the change previewed under `token`. This performs a REAL CHANGE to the "
    "user's calendar and cannot be undone by this server. The token must come from "
    "preview_event_update in this same session; each token works exactly once, and an "
    "unknown or already used token is refused, as is a token issued by preview_event or "
    "preview_event_deletion. Call this only after the user has confirmed the preview's "
    "summary. " + CONFIRMATION_RULE
)


def register_calendar_update_tools(server: MCPServer, use_cases: CalendarWriteUseCases) -> None:
    _register_preview_event_update(server, use_cases.preview_event_update)
    _register_update_event(server, use_cases.update_event)


def _register_preview_event_update(server: MCPServer, use_case: PreviewEventUpdate) -> None:
    @server.tool(name=PREVIEW_EVENT_UPDATE_TOOL, description=PREVIEW_EVENT_UPDATE_DESCRIPTION)
    async def preview_event_update(
        event_id: EventId,
        subject: Subject | None = None,
        start: Start | None = None,
        end: End | None = None,
        location: Location | None = None,
        body: Body | None = None,
        attendees: Attendees | None = None,
        reminder_minutes_before_start: ReminderMinutesBeforeStart | None = None,
        sensitivity: EventSensitivity | None = None,
        show_as: EventShowAs | None = None,
    ) -> EventUpdateDraftView:
        model = UpdateEventInput(
            event_id=event_id,
            subject=subject,
            start=start,
            end=end,
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


def _translate_preview(
    use_case: PreviewEventUpdate, model: UpdateEventInput
) -> EventUpdateDraftView:
    with observed_tool_call(PREVIEW_EVENT_UPDATE_TOOL) as outcome:
        draft = use_case.execute(model.to_changes())
        outcome.item_count = 1
        return EventUpdateDraftView.from_draft(draft)


def _register_update_event(server: MCPServer, use_case: UpdateEvent) -> None:
    @server.tool(name=UPDATE_EVENT_TOOL, description=UPDATE_EVENT_DESCRIPTION)
    async def update_event(token: str) -> EventDetailView:
        try:
            return _translate_update(use_case, token)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_update(use_case: UpdateEvent, token: str) -> EventDetailView:
    with observed_tool_call(UPDATE_EVENT_TOOL) as outcome:
        event = use_case.execute(token)
        outcome.item_count = 1
        return EventDetailView.from_event(event)
