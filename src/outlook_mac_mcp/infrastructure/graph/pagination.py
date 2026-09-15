from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError

NEXT_LINK_FIELD = "@odata.nextLink"


def read_next_link(payload: Mapping[str, Any]) -> str | None:
    """Return the link to the next page, or None when this page is the last one."""
    next_link = payload.get(NEXT_LINK_FIELD)
    if next_link is None:
        return None
    if not isinstance(next_link, str) or not next_link:
        raise GraphResponseError(f"the message collection carried a malformed {NEXT_LINK_FIELD}")
    return next_link
