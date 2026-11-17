from pydantic import BaseModel, ConfigDict

from outlook_mac_mcp.application.preview_draft_deletion import DraftDeletionRequest
from outlook_mac_mcp.interface.mcp.get_email_input import EmailId


class DraftDeletionInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    email_id: EmailId

    def to_request(self) -> DraftDeletionRequest:
        return DraftDeletionRequest(email_id=self.email_id)
