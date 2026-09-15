from outlook_mac_mcp.application.search_emails_request import (
    UNSUPPORTED_TERM_CHARACTERS,
)
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.search_scope import SearchScope

# The property names Graph's KQL uses for each scope. They come from this table and never
# from the caller, which is what keeps a scoped search free of injected restrictions.
KQL_PROPERTY_BY_SCOPE = {
    SearchScope.SUBJECT: "subject",
    SearchScope.SENDER: "from",
}


def to_search_query(term: str, scope: SearchScope) -> str:
    """Build the $search value: a nested KQL phrase, optionally restricted to a property.

    Graph reads the whole $search value as one OData string and parses KQL *inside* it,
    so the outer quotes are the string delimiter, not a phrase delimiter. Sending
    `"deck AND from:x"` therefore executes the operators; the term is only literal inside
    its own escaped phrase, `"\\"deck AND from:x\\""`. Verified against the live API: the
    first form returns the filtered result, the second returns none.

    A property restriction goes inside the same string, as `"subject:\\"deck\\""`. Placing
    it outside, as `subject:"deck"`, is rejected by Graph with 400.

    The term is guarded here as well as in the request, because this is the last point
    before the wire and the quoting above is only sound for a term that cannot close it.
    """
    _reject_unquotable(term)
    phrase = f'\\"{term}\\"'
    if scope is SearchScope.ANY:
        return f'"{phrase}"'
    return f'"{KQL_PROPERTY_BY_SCOPE[scope]}:{phrase}"'


def _reject_unquotable(term: str) -> None:
    if any(character in term for character in UNSUPPORTED_TERM_CHARACTERS):
        raise InvalidRequestError("term must not contain a double quote or a backslash")
