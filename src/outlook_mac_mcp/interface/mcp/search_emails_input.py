from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.search_emails_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchEmailsRequest,
)
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.interface.mcp.email_filters_input import MAX_FOLDER_ARGUMENT_LENGTH

# Mirrors UNSUPPORTED_TERM_CHARACTERS so the refusal appears in the tool schema rather
# than only as a runtime error.
TERM_PATTERN = r'^[^"\\]*$'

Term = Annotated[
    str,
    Field(
        min_length=MIN_TERM_LENGTH,
        max_length=MAX_TERM_LENGTH,
        pattern=TERM_PATTERN,
        description=(
            "Text to look for. Matched literally: query operators such as from: or AND "
            "are searched for as words, not interpreted. Double quotes and backslashes "
            "are not accepted."
        ),
    ),
]
SearchFolder = Annotated[
    str,
    Field(
        min_length=1,
        max_length=MAX_FOLDER_ARGUMENT_LENGTH,
        description=(
            "Mailbox folder to search: a well-known name (default inbox), all (the five "
            "well-known folders only -- never custom folders), or a custom folder's "
            "full path as shown in list_folders' custom list, e.g. Entrevistas/Work/AWS."
        ),
    ),
]
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
PageToken = Annotated[
    str,
    Field(
        min_length=1,
        description=(
            "Continuation token from a previous search_emails call's next_page_token, "
            "to fetch the next page of that same search. Omit to start a new search "
            "from the first page."
        ),
    ),
]


class SearchEmailsInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    term: Term
    folder: SearchFolder = FolderSelection.INBOX
    scope: Scope = SearchScope.ANY
    limit: SearchLimit = DEFAULT_LIMIT
    page_token: PageToken | None = None

    def to_request(self, folder_ids: tuple[str, ...]) -> SearchEmailsRequest:
        return SearchEmailsRequest(
            term=self.term,
            folders=folder_ids,
            scope=self.scope,
            limit=self.limit,
            page_token=self.page_token,
        )
