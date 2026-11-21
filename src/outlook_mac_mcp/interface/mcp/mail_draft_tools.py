from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mac_mcp.application.create_mail_draft import CreateMailDraft
from outlook_mac_mcp.application.preview_mail_draft import PreviewMailDraft
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.interface.mcp.created_mail_draft_view import CreatedMailDraftView
from outlook_mac_mcp.interface.mcp.draft_deletion_tools import register_draft_deletion_tools
from outlook_mac_mcp.interface.mcp.mail_draft_preview_view import MailDraftPreviewView
from outlook_mac_mcp.interface.mcp.mail_draft_use_cases import MailDraftUseCases
from outlook_mac_mcp.interface.mcp.observability import observed_tool_call
from outlook_mac_mcp.interface.mcp.preview_draft_input import (
    CcRecipients,
    DraftBody,
    DraftSubject,
    PreviewDraftInput,
    ToRecipients,
)
from outlook_mac_mcp.interface.mcp.reply_draft_tools import register_reply_draft_tools
from outlook_mac_mcp.interface.mcp.write_confirmation_rule import write_confirmation_rule

PREVIEW_DRAFT_TOOL = "preview_draft"
CREATE_DRAFT_TOOL = "create_draft"
CONFIRMATION_RULE = write_confirmation_rule(PREVIEW_DRAFT_TOOL)
PREVIEW_DRAFT_DESCRIPTION = (
    "Prepare an email draft without saving it. Validates the details and returns a token, "
    "a one-line summary and any warnings, such as each link in the body. Nothing is written "
    "to the mailbox by this tool. Show the summary and EVERY warning to the user and, only "
    "once they agree, pass the token to create_draft. " + CONFIRMATION_RULE
)
CREATE_DRAFT_DESCRIPTION = (
    "Save the draft previewed under `token` in the user's Drafts folder. This is a REAL "
    "CHANGE to the mailbox, but it only saves: the message is NOT sent, and this server "
    "has no way to send it. The user reviews and sends it from Outlook. The token must come "
    "from preview_draft in this same session; each token works exactly once, and an unknown "
    "or already used token is refused. Call this only after the user has confirmed the "
    "preview's summary. " + CONFIRMATION_RULE
)


def register_mail_draft_tools(server: MCPServer, use_cases: MailDraftUseCases) -> None:
    _register_preview_draft(server, use_cases.preview_mail_draft)
    _register_create_draft(server, use_cases.create_mail_draft)
    register_reply_draft_tools(server, use_cases.preview_reply_draft, use_cases.create_reply_draft)
    if use_cases.preview_draft_deletion is not None and use_cases.delete_draft is not None:
        register_draft_deletion_tools(
            server, use_cases.preview_draft_deletion, use_cases.delete_draft
        )


def _register_preview_draft(server: MCPServer, use_case: PreviewMailDraft) -> None:
    @server.tool(name=PREVIEW_DRAFT_TOOL, description=PREVIEW_DRAFT_DESCRIPTION)
    async def preview_draft(
        subject: DraftSubject,
        body: DraftBody,
        to: ToRecipients,
        cc: CcRecipients = (),
    ) -> MailDraftPreviewView:
        model = PreviewDraftInput(subject=subject, body=body, to=to, cc=cc)
        try:
            return _translate_preview(use_case, model)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_preview(
    use_case: PreviewMailDraft, model: PreviewDraftInput
) -> MailDraftPreviewView:
    with observed_tool_call(PREVIEW_DRAFT_TOOL) as outcome:
        preview = use_case.execute(model.to_new_mail_draft())
        outcome.item_count = 1
        return MailDraftPreviewView.from_preview(preview)


def _register_create_draft(server: MCPServer, use_case: CreateMailDraft) -> None:
    @server.tool(name=CREATE_DRAFT_TOOL, description=CREATE_DRAFT_DESCRIPTION)
    async def create_draft(token: str) -> CreatedMailDraftView:
        try:
            return _translate_create(use_case, token)
        except OutlookMcpError as error:
            raise ToolError(str(error)) from error


def _translate_create(use_case: CreateMailDraft, token: str) -> CreatedMailDraftView:
    with observed_tool_call(CREATE_DRAFT_TOOL) as outcome:
        created = use_case.execute(token)
        outcome.item_count = 1
        return CreatedMailDraftView.from_created(created)
