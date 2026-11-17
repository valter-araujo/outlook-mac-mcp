from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.create_reply_draft import CreateReplyDraft
from outlook_mac_mcp.application.preview_reply_draft import PreviewReplyDraft
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.interface.mcp.created_mail_draft_view import CreatedMailDraftView
from outlook_mac_mcp.interface.mcp.get_email_input import EmailId
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.reply_draft_input import (
    ExtraCc,
    ExtraTo,
    Mode,
    PreviewReplyDraftInput,
    ReplyBody,
)
from outlook_mac_mcp.interface.mcp.reply_draft_preview_view import ReplyDraftPreviewView
from outlook_mac_mcp.interface.mcp.write_confirmation_rule import write_confirmation_rule

PREVIEW_REPLY_DRAFT_TOOL = "preview_reply_draft"
CREATE_REPLY_DRAFT_TOOL = "create_reply_draft"
CONFIRMATION_RULE = write_confirmation_rule(PREVIEW_REPLY_DRAFT_TOOL)
PREVIEW_REPLY_DRAFT_DESCRIPTION = (
    "Prepare a reply or reply-all draft to an existing email without saving it. Use this, "
    "not preview_draft, whenever the user wants to answer a message: the reply stays in the "
    "original conversation and keeps its recipients. Reads the original and returns a token, "
    "a summary, the final To and Cc lines (Reply-To honoured, the user's own address left "
    "out) and any warnings, such as a recipient that looks like a no-reply address. Nothing "
    "is written to the mailbox by this tool. Show the user the summary, EVERY recipient and "
    "EVERY warning, and only once they agree pass the token to create_reply_draft. "
    + CONFIRMATION_RULE
)
CREATE_REPLY_DRAFT_DESCRIPTION = (
    "Save the reply previewed under `token` as a draft in the original email's conversation, "
    "with the original quoted below the new text. This is a REAL CHANGE to the mailbox, but "
    "it only saves: the reply is NOT sent, and this server has no way to send it. The user "
    "reviews and sends it from Outlook. The token must come from preview_reply_draft in this "
    "same session; each token works exactly once, and an unknown or already used token is "
    "refused. Call this only after the user has confirmed the preview. " + CONFIRMATION_RULE
)


def register_reply_draft_tools(
    server: MCPServer, preview: PreviewReplyDraft, create: CreateReplyDraft
) -> None:
    _register_preview_reply_draft(server, preview)
    _register_create_reply_draft(server, create)


def _register_preview_reply_draft(server: MCPServer, use_case: PreviewReplyDraft) -> None:
    @server.tool(name=PREVIEW_REPLY_DRAFT_TOOL, description=PREVIEW_REPLY_DRAFT_DESCRIPTION)
    async def preview_reply_draft(
        email_id: EmailId,
        body: ReplyBody,
        mode: Mode = ReplyMode.REPLY_ALL,
        extra_to: ExtraTo = (),
        extra_cc: ExtraCc = (),
    ) -> ReplyDraftPreviewView:
        model = PreviewReplyDraftInput(
            email_id=email_id, mode=mode, body=body, extra_to=extra_to, extra_cc=extra_cc
        )
        try:
            return _translate_preview(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_preview(
    use_case: PreviewReplyDraft, model: PreviewReplyDraftInput
) -> ReplyDraftPreviewView:
    with observed_tool_call(PREVIEW_REPLY_DRAFT_TOOL) as outcome:
        preview = use_case.execute(model.to_request())
        outcome.item_count = 1
        return ReplyDraftPreviewView.from_preview(preview)


def _register_create_reply_draft(server: MCPServer, use_case: CreateReplyDraft) -> None:
    @server.tool(name=CREATE_REPLY_DRAFT_TOOL, description=CREATE_REPLY_DRAFT_DESCRIPTION)
    async def create_reply_draft(token: str) -> CreatedMailDraftView:
        try:
            return _translate_create(use_case, token)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_create(use_case: CreateReplyDraft, token: str) -> CreatedMailDraftView:
    with observed_tool_call(CREATE_REPLY_DRAFT_TOOL) as outcome:
        created = use_case.execute(token)
        outcome.item_count = 1
        return CreatedMailDraftView.from_created(created)
