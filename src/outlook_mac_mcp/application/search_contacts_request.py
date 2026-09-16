from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.domain.errors import InvalidRequestError

MIN_TERM_LENGTH = 1
MAX_TERM_LENGTH = 200

# The term is spliced into two OData string literals (a startswith and an eq); a single
# quote is that literal's delimiter, so refusing it is simpler and safer than escaping
# it, the same choice EmailFilters.sender makes for the same reason.
UNSUPPORTED_TERM_CHARACTERS = "'"


@dataclass(frozen=True, slots=True)
class SearchContactsRequest:
    """A validated search, and the single argument the repository port takes."""

    term: str
    limit: int = DEFAULT_LIMIT

    def __post_init__(self) -> None:
        if not MIN_TERM_LENGTH <= len(self.term) <= MAX_TERM_LENGTH:
            raise InvalidRequestError(
                f"term must be between {MIN_TERM_LENGTH} and {MAX_TERM_LENGTH} characters"
            )
        if any(character in self.term for character in UNSUPPORTED_TERM_CHARACTERS):
            raise InvalidRequestError("term must not contain a single quote")
        ensure_limit_within_bounds(self.limit)
