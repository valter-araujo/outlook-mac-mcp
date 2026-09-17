from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.folder_selection import ensure_well_formed_folders
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

    `folders` holds more than one folder only for the `all` selection: the adapter
    searches each and merges the results, still ranked by relevance within each folder's
    own results, not globally re-ranked across folders (Graph exposes no cross-folder
    relevance score to merge by).

    `page_token` is opaque to everything above the Graph adapter: None means the first
    page of a new search, built from `term`/`folders`/`scope`; anything else is a
    continuation token from a previous page, tied to whichever one folder it was issued
    for, and the adapter follows it directly instead of rebuilding the query. The total
    is still recomputed across every folder in `folders`, so it keeps meaning "how many
    match this search in all", not "in the one folder this page happens to come from".
    """

    term: str
    folders: tuple[FolderName, ...] = (FolderName.INBOX,)
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
        ensure_well_formed_folders(self.folders)
