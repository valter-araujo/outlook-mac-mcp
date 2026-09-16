"""Building the $filter behind search_contacts, and why the email side is exact.

Graph's List contacts documentation is explicit about the restriction:

    "You can use $filter, any, and the eq operator on only the address sub-property of
    instances in an emailAddresses collection. That is, you can't ... apply any other
    operator or function with filter, such as ne, le, and startswith()."
    https://learn.microsoft.com/en-us/graph/api/user-list-contacts

So a prefix match on email address, as this project's other filters allow, is not an
option Graph accepts: the email side of the search can only be an exact address, while
the display-name side keeps the ordinary startswith a name search calls for. A term that
is not a complete address therefore matches only by name; the tool description says so.

Untested against a live mailbox as of this writing; see the project's rule in CLAUDE.md
that a property depending on how an external API parses input needs a live positive-
control check before it can be trusted, the same rule that caught the $search quoting
issue in search_query.py.
"""

from outlook_mac_mcp.application.search_contacts_request import UNSUPPORTED_TERM_CHARACTERS
from outlook_mac_mcp.domain.errors import InvalidRequestError

NAME_PROPERTY = "displayName"
EMAIL_COLLECTION = "emailAddresses"


def contact_search_query(term: str) -> dict[str, str]:
    _reject_unquotable(term)
    name_clause = f"startswith({NAME_PROPERTY},'{term}')"
    email_clause = f"{EMAIL_COLLECTION}/any(a:a/address eq '{term}')"
    return {"$filter": f"{name_clause} or {email_clause}", "$count": "true"}


def _reject_unquotable(term: str) -> None:
    """Last point before the wire: a single quote would close the OData literal early."""
    if any(character in term for character in UNSUPPORTED_TERM_CHARACTERS):
        raise InvalidRequestError("term must not contain a single quote")
