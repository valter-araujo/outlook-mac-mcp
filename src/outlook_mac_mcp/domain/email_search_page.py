from dataclasses import dataclass

from outlook_mac_mcp.domain.email import Email


@dataclass(frozen=True, slots=True)
class EmailSearchPage:
    """One page of a search_emails result, together with an opaque continuation token.

    Distinct from Page because search is the only listing this project paginates:
    Graph's $search carries `@odata.nextLink` when more matches exist beyond this page,
    and `next_page_token` exposes it verbatim so a caller can ask for the next page of
    the same search without knowing anything about Graph's pagination mechanics. None
    means this page is the last one.
    """

    items: tuple[Email, ...]
    total: int
    total_is_exact: bool
    next_page_token: str | None
