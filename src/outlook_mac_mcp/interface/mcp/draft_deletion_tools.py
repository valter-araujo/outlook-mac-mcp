from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.delete_draft import DeleteDraft
from outlook_mac_mcp.application.preview_draft_deletion import PreviewDraftDeletion
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.draft_deletion_input import DraftDeletionInput
from outlook_mac_mcp.interface.mcp.draft_deletion_preview_view import DraftDeletionPreviewView
from outlook_mac_mcp.interface.mcp.draft_deletion_result_view import DraftDeletionResultView
from outlook_mac_mcp.interface.mcp.get_email_input import EmailId
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call

PREVIEW_DRAFT_DELETION_TOOL = "preview_draft_deletion"
DELETE_DRAFT_TOOL = "delete_draft"
PREVIEW_DRAFT_DELETION_DESCRIPTION = (
    "Prepare the deletion of an email draft without deleting it, typically to clean up a "
    "draft saved by mistake. Only an unsent draft in the Drafts folder is accepted; any "
    "other message is refused. Returns a token and a summary (subject, To line, last edit "
    "time). Nothing is deleted by this tool. Show the summary to the user and, only once "
    "they agree, pass the token to delete_draft."
)
DELETE_DRAFT_DESCRIPTION = (
    "Delete the draft previewed under `token`. This is a REAL CHANGE to the mailbox: the "
    "draft moves to Deleted Items, where the user can still recover it. The draft is "
    "checked again first and the deletion is refused if it was sent or moved out of Drafts "
    "since the preview. The token must come from preview_draft_deletion in this same "
    "session; each token works exactly once, and an unknown or already used token is "
    "refused. Call this only after the user has confirmed the preview's summary."
)


def register_draft_deletion_tools(
    server: MCPServer, preview: PreviewDraftDeletion, delete: DeleteDraft
) -> None:
    _register_preview_draft_deletion(server, preview)
    _register_delete_draft(server, delete)


def _register_preview_draft_deletion(server: MCPServer, use_case: PreviewDraftDeletion) -> None:
    @server.tool(name=PREVIEW_DRAFT_DELETION_TOOL, description=PREVIEW_DRAFT_DELETION_DESCRIPTION)
    async def preview_draft_deletion(email_id: EmailId) -> DraftDeletionPreviewView:
        model = DraftDeletionInput(email_id=email_id)
        try:
            return _translate_preview(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_preview(
    use_case: PreviewDraftDeletion, model: DraftDeletionInput
) -> DraftDeletionPreviewView:
    with observed_tool_call(PREVIEW_DRAFT_DELETION_TOOL) as outcome:
        preview = use_case.execute(model.to_request())
        outcome.item_count = 1
        return DraftDeletionPreviewView.from_preview(preview)


def _register_delete_draft(server: MCPServer, use_case: DeleteDraft) -> None:
    @server.tool(name=DELETE_DRAFT_TOOL, description=DELETE_DRAFT_DESCRIPTION)
    async def delete_draft(token: str) -> DraftDeletionResultView:
        try:
            return _translate_delete(use_case, token)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_delete(use_case: DeleteDraft, token: str) -> DraftDeletionResultView:
    with observed_tool_call(DELETE_DRAFT_TOOL) as outcome:
        draft_id = use_case.execute(token)
        outcome.item_count = 1
        return DraftDeletionResultView.from_draft_id(draft_id)
