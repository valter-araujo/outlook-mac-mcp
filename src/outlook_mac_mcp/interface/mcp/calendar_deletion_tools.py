from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.delete_event import DeleteEvent
from outlook_mac_mcp.application.preview_event_deletion import PreviewEventDeletion
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.calendar_write_use_cases import CalendarWriteUseCases
from outlook_mac_mcp.interface.mcp.event_deletion_draft_view import EventDeletionDraftView
from outlook_mac_mcp.interface.mcp.event_deletion_input import EventDeletionInput, EventId
from outlook_mac_mcp.interface.mcp.event_deletion_result_view import EventDeletionResultView
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.write_confirmation_rule import write_confirmation_rule

PREVIEW_EVENT_DELETION_TOOL = "preview_event_deletion"
DELETE_EVENT_TOOL = "delete_event"
CONFIRMATION_RULE = write_confirmation_rule(PREVIEW_EVENT_DELETION_TOOL)
PREVIEW_EVENT_DELETION_DESCRIPTION = (
    "Prepare the deletion of an existing event without deleting it. Fetches the event "
    "and returns a token and a summary showing every field of the event as it "
    "currently stands, unabbreviated, so the confirmation shows exactly what would be "
    "removed. This matters most when two similar events exist and only their full "
    "detail tells them apart: check the summary carefully before proceeding. Nothing "
    "is deleted by this tool. Show the summary to the user and, only once they agree, "
    "pass the token to delete_event. " + CONFIRMATION_RULE
)
DELETE_EVENT_DESCRIPTION = (
    "Delete the event previewed under `token`. This performs a REAL, IRREVERSIBLE "
    "CHANGE to the user's calendar: the event cannot be recovered by this server. The "
    "token must come from preview_event_deletion in this same session; each token "
    "works exactly once, and an unknown or already used token is refused, as is a "
    "token issued by preview_event or preview_event_update. Call this only after the "
    "user has confirmed the preview's summary. " + CONFIRMATION_RULE
)


def register_calendar_deletion_tools(server: MCPServer, use_cases: CalendarWriteUseCases) -> None:
    _register_preview_event_deletion(server, use_cases.preview_event_deletion)
    _register_delete_event(server, use_cases.delete_event)


def _register_preview_event_deletion(server: MCPServer, use_case: PreviewEventDeletion) -> None:
    @server.tool(name=PREVIEW_EVENT_DELETION_TOOL, description=PREVIEW_EVENT_DELETION_DESCRIPTION)
    async def preview_event_deletion(event_id: EventId) -> EventDeletionDraftView:
        model = EventDeletionInput(event_id=event_id)
        try:
            return _translate_preview(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_preview(
    use_case: PreviewEventDeletion, model: EventDeletionInput
) -> EventDeletionDraftView:
    with observed_tool_call(PREVIEW_EVENT_DELETION_TOOL) as outcome:
        draft = use_case.execute(model.to_request())
        outcome.item_count = 1
        return EventDeletionDraftView.from_draft(draft)


def _register_delete_event(server: MCPServer, use_case: DeleteEvent) -> None:
    @server.tool(name=DELETE_EVENT_TOOL, description=DELETE_EVENT_DESCRIPTION)
    async def delete_event(token: str) -> EventDeletionResultView:
        try:
            return _translate_delete(use_case, token)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_delete(use_case: DeleteEvent, token: str) -> EventDeletionResultView:
    with observed_tool_call(DELETE_EVENT_TOOL) as outcome:
        event_id = use_case.execute(token)
        outcome.item_count = 1
        return EventDeletionResultView.from_event_id(event_id)
