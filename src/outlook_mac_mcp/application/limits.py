"""How many items a read use case may return.

One policy shared by every listing use case, so a caller sees the same bound whichever
tool it reaches for, and the interface layer derives its schema from a single source.
"""

from outlook_mac_mcp.domain.errors import InvalidRequestError

MIN_LIMIT = 1
MAX_LIMIT = 100
DEFAULT_LIMIT = 20


def ensure_limit_within_bounds(limit: int) -> None:
    if not MIN_LIMIT <= limit <= MAX_LIMIT:
        raise InvalidRequestError(f"limit must be between {MIN_LIMIT} and {MAX_LIMIT}")
