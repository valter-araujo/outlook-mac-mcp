from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.search_scope import SearchScope

MIN_TERM_LENGTH = 1
MAX_TERM_LENGTH = 200


@dataclass(frozen=True, slots=True)
class SearchEmailsRequest:
    """A validated search, and the single argument the repository port takes.

    It lives in its own module so the port can name it without importing the use case
    that consumes it.
    """

    term: str
    folder: FolderName = FolderName.INBOX
    scope: SearchScope = SearchScope.ANY
    limit: int = DEFAULT_LIMIT

    def __post_init__(self) -> None:
        if not MIN_TERM_LENGTH <= len(self.term) <= MAX_TERM_LENGTH:
            raise InvalidRequestError(
                f"term must be between {MIN_TERM_LENGTH} and {MAX_TERM_LENGTH} characters"
            )
        ensure_limit_within_bounds(self.limit)
