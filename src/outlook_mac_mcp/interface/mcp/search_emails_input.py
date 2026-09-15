from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.search_emails_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchEmailsRequest,
)
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.search_scope import SearchScope

Term = Annotated[
    str,
    Field(
        min_length=MIN_TERM_LENGTH,
        max_length=MAX_TERM_LENGTH,
        description=(
            "Text to look for. Matched literally: query operators such as from: or AND "
            "are searched for as words, not interpreted."
        ),
    ),
]
SearchFolder = Annotated[FolderName, Field(description="Mailbox folder to search.")]
Scope = Annotated[
    SearchScope,
    Field(
        description=(
            "Which part of the email to match: any searches subject, body and sender; "
            "subject matches the subject line only; sender matches the sender only."
        )
    ),
]
SearchLimit = Annotated[
    int,
    Field(ge=MIN_LIMIT, le=MAX_LIMIT, description="Maximum number of emails to return."),
]


class SearchEmailsInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    term: Term
    folder: SearchFolder = FolderName.INBOX
    scope: Scope = SearchScope.ANY
    limit: SearchLimit = DEFAULT_LIMIT

    def to_request(self) -> SearchEmailsRequest:
        return SearchEmailsRequest(
            term=self.term, folder=self.folder, scope=self.scope, limit=self.limit
        )
