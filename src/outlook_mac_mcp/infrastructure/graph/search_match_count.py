"""Counting the matches of a Graph $search, which $count=true cannot do.

Graph rejects `$count` alongside `$search`, so the only way to learn how many messages
match is to walk the result pages. Walking them all would cost one request per page for
a term that matches thousands of messages, so the walk stops at a ceiling: below it the
count is exact, at it the count is a lower bound and the caller is told so.
"""

from dataclasses import dataclass

from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link

COUNT_CEILING = 250
ID_ONLY_SELECT = "id"


@dataclass(frozen=True, slots=True)
class MatchCount:
    total: int
    is_exact: bool


def count_search_matches(client: GraphClient, path: str, search: str) -> MatchCount:
    """Walk the pages of `search` under `path`, fetching ids only, until the ceiling.

    The first page asks for the whole ceiling at once; Graph may still answer in smaller
    pages, which is why the links are followed rather than the first page trusted.
    """
    payload = client.get(
        path, {"$search": search, "$top": COUNT_CEILING, "$select": ID_ONLY_SELECT}
    )
    counted = 0
    while True:
        counted += len(read_items(payload))
        next_link = read_next_link(payload)
        if counted > COUNT_CEILING or (counted == COUNT_CEILING and next_link is not None):
            return MatchCount(total=COUNT_CEILING, is_exact=False)
        if next_link is None:
            return MatchCount(total=counted, is_exact=True)
        payload = client.follow(next_link)
