from outlook_mac_mcp.domain.search_scope import SearchScope

BACKSLASH = "\\"
DOUBLE_QUOTE = '"'


def to_literal_phrase(term: str) -> str:
    """Turn a search term into a KQL phrase that cannot carry operators.

    Graph parses `$search` as KQL, where bare text such as `from:someone`, `AND` or `NOT`
    are operators rather than words. Wrapping the whole term in double quotes makes it a
    single literal phrase, so a term that looks like a query is searched for, not obeyed.

    Backslashes are escaped before quotes: escaping in the other order would let a term
    ending in a backslash consume the escape and close the phrase early.
    """
    escaped = term.replace(BACKSLASH, BACKSLASH * 2).replace(DOUBLE_QUOTE, BACKSLASH + DOUBLE_QUOTE)
    return f"{DOUBLE_QUOTE}{escaped}{DOUBLE_QUOTE}"


# The property names Graph's KQL uses for each scope. They come from this table and never
# from the caller, which is what keeps a scoped search free of injected restrictions.
KQL_PROPERTY_BY_SCOPE = {
    SearchScope.SUBJECT: "subject",
    SearchScope.SENDER: "from",
}


def to_search_query(term: str, scope: SearchScope) -> str:
    """Build the $search value: a quoted phrase, optionally restricted to one property.

    `subject:"deck"` narrows the match to the subject line, where an unscoped search also
    reads the body and turns any newsletter that mentions the term into a hit.
    """
    phrase = to_literal_phrase(term)
    if scope is SearchScope.ANY:
        return phrase
    return f"{KQL_PROPERTY_BY_SCOPE[scope]}:{phrase}"
