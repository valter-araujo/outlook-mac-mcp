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
