from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError

NEXT_LINK_FIELD = "@odata.nextLink"


def read_items(payload: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    """Return the page's `value` array, each entry checked to be an object."""
    items = payload.get("value")
    if not isinstance(items, list):
        raise GraphResponseError("the collection carried no value array")
    for item in items:
        if not isinstance(item, dict):
            raise GraphResponseError("the collection held something other than an object")
    return items


def read_next_link(payload: Mapping[str, Any]) -> str | None:
    """Return the link to the next page, or None when this page is the last one."""
    next_link = payload.get(NEXT_LINK_FIELD)
    if next_link is None:
        return None
    if not isinstance(next_link, str) or not next_link:
        raise GraphResponseError(f"the collection carried a malformed {NEXT_LINK_FIELD}")
    return next_link
