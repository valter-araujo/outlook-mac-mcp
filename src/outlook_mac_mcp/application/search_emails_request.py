from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.search_scope import SearchScope

MIN_TERM_LENGTH = 1
MAX_TERM_LENGTH = 200

# A double quote or a backslash cannot be carried safely into a Graph KQL phrase:
# escaping them is not parsed reliably, and a term ending in a backslash was observed
# to break the phrase open and match the entire mailbox. Refusing the two characters
# costs a rare search and removes the escaping problem altogether.
UNSUPPORTED_TERM_CHARACTERS = '"\\'


@dataclass(frozen=True, slots=True)
class SearchEmailsRequest:
    """A validated search, and the single argument the repository port takes.

    It lives in its own module so the port can name it without importing the use case
    that consumes it.

    `page_token` is opaque to everything above the Graph adapter: None means the first
    page of a new search, built from `term`/`folder`/`scope`; anything else is a
    continuation token from a previous page, and the adapter follows it directly instead
    of rebuilding the query. `term`/`folder`/`scope` are still validated and still
    present on a continuation call, but the adapter does not use them to fetch the page.
    """

    term: str
    folder: FolderName = FolderName.INBOX
    scope: SearchScope = SearchScope.ANY
    limit: int = DEFAULT_LIMIT
    page_token: str | None = None

    def __post_init__(self) -> None:
        if not MIN_TERM_LENGTH <= len(self.term) <= MAX_TERM_LENGTH:
            raise InvalidRequestError(
                f"term must be between {MIN_TERM_LENGTH} and {MAX_TERM_LENGTH} characters"
            )
        if any(character in self.term for character in UNSUPPORTED_TERM_CHARACTERS):
            raise InvalidRequestError("term must not contain a double quote or a backslash")
        ensure_limit_within_bounds(self.limit)
