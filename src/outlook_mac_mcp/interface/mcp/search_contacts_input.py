from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.search_contacts_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchContactsRequest,
)

# Mirrors UNSUPPORTED_TERM_CHARACTERS so the refusal appears in the tool schema.
TERM_PATTERN = r"^[^']*$"

Term = Annotated[
    str,
    Field(
        min_length=MIN_TERM_LENGTH,
        max_length=MAX_TERM_LENGTH,
        pattern=TERM_PATTERN,
        description=(
            "Text to match. Matched as a prefix against the display name, and as an "
            "exact address against any of the contact's email addresses. A single quote "
            "is not accepted."
        ),
    ),
]
Limit = Annotated[
    int,
    Field(ge=MIN_LIMIT, le=MAX_LIMIT, description="Maximum number of contacts to return."),
]


class SearchContactsInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    term: Term
    limit: Limit = DEFAULT_LIMIT

    def to_request(self) -> SearchContactsRequest:
        return SearchContactsRequest(term=self.term, limit=self.limit)
