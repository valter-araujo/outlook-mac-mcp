r"""Building the Graph $search value, and why it is shaped the way it is.

Live check against Microsoft Graph on 2026-09-14, personal Outlook.com mailbox, inbox,
$top=100. The positive control is `from:<a real sender address>`, which matches exactly
one message: a term carrying that operator returns one result if the operator ran, and
none if the term was read as text.

    "<term>"                                        matched several messages
    "<term> AND from:<sender>"                       filtered to the one control message,
                                                      operators executed
    "\"<term> AND from:<sender>\""                   matched nothing, literal, as intended
    subject:"<term>"                                 400      BadRequest
    "subject:\"<term>\""                             matched a narrower, still-exact subset
    "\"path\\\""                                     matched a full page, phrase boundary lost

Three conclusions, none of which the unit tests could have reached, because asserting
the bytes we send says nothing about how Graph parses them:

1. The whole $search value is one OData string with KQL parsed inside it, so the outer
   quotes are the string delimiter and not a phrase delimiter. A singly quoted term has
   its operators executed.
2. The term is literal only inside its own nested phrase, and a property restriction
   belongs inside that same string. Placing the property outside is rejected with 400.
3. Escaping cannot be relied on to keep the phrase closed: a term ending in a backslash
   lost the phrase boundary and matched the whole mailbox, returning a full page. Hence
   the refusal in `_reject_unquotable` rather than an escape.
"""

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
    r"""Build the $search value: a nested KQL phrase, optionally restricted to a property.

    The outer quotes are the OData string; the inner \" pair is the KQL phrase that makes
    the term literal. A property restriction goes inside the same string. See the module
    docstring for the measurements behind each of those.

    The term is guarded here as well as in the request because this is the last point
    before the wire, and the quoting is only sound for a term that cannot close it.
    """
    _reject_unquotable(term)
    phrase = f'\\"{term}\\"'
    if scope is SearchScope.ANY:
        return f'"{phrase}"'
    return f'"{KQL_PROPERTY_BY_SCOPE[scope]}:{phrase}"'


def _reject_unquotable(term: str) -> None:
    """Refuse rather than escape: a trailing backslash loses the phrase boundary.

    A term ending in a backslash came back with a full page of results against the live
    API, meaning the phrase never closed and the query matched the whole mailbox. A
    quoting scheme that can lose its own boundary cannot carry a security property.
    """
    if any(character in term for character in UNSUPPORTED_TERM_CHARACTERS):
        raise InvalidRequestError("term must not contain a double quote or a backslash")
