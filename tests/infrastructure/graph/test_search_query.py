import pytest

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.infrastructure.graph.search_query import to_search_query

# `"deck"` on the wire is the OData string delimiter; `\"deck\"` inside it is the KQL
# phrase. Only the nested form is literal, which the live API confirmed.
OPERATOR_LIKE_TERMS = [
    "from:ceo@example.com",
    "subject:payroll",
    "AND",
    "OR",
    "NOT",
    "deck AND from:ana@example.com",
    "received>=2026-01-01",
    "(a OR b)",
    "ratio 3:1",
]


def test_an_any_scope_wraps_the_term_in_a_nested_phrase() -> None:
    assert to_search_query("Cargill", SearchScope.ANY) == '"\\"Cargill\\""'


def test_a_subject_scope_puts_the_property_inside_the_string() -> None:
    """Graph rejects `subject:"Cargill"` with 400; the property belongs inside the quotes."""
    assert to_search_query("Cargill", SearchScope.SUBJECT) == '"subject:\\"Cargill\\""'


def test_a_sender_scope_uses_the_from_property() -> None:
    """Graph's KQL calls it `from`; the enum name is ours, the property name is theirs."""
    assert to_search_query("Cargill", SearchScope.SENDER) == '"from:\\"Cargill\\""'


@pytest.mark.parametrize("scope", list(SearchScope))
def test_every_query_is_one_odata_string(scope: SearchScope) -> None:
    query = to_search_query("Cargill", scope)

    assert query.startswith('"')
    assert query.endswith('"')


@pytest.mark.parametrize("term", OPERATOR_LIKE_TERMS)
@pytest.mark.parametrize("scope", list(SearchScope))
def test_text_that_reads_like_a_query_stays_inside_the_phrase(
    term: str, scope: SearchScope
) -> None:
    query = to_search_query(term, scope)

    assert query.endswith(f'\\"{term}\\""')


def test_a_term_that_names_a_property_is_not_promoted_to_one() -> None:
    assert to_search_query("subject:payroll", SearchScope.SENDER) == (
        '"from:\\"subject:payroll\\""'
    )


@pytest.mark.parametrize("scope", list(SearchScope))
def test_rejects_a_term_holding_a_double_quote(scope: SearchScope) -> None:
    """The quote is what a breakout needs, and escaping it is not parsed reliably."""
    with pytest.raises(InvalidRequestError):
        to_search_query('deck" AND from:ceo@example.com "', scope)


@pytest.mark.parametrize("scope", list(SearchScope))
def test_rejects_a_term_holding_a_backslash(scope: SearchScope) -> None:
    """A trailing backslash was observed to break the phrase open and match everything."""
    with pytest.raises(InvalidRequestError):
        to_search_query("path\\", scope)


def test_accepts_a_term_with_an_ordinary_apostrophe() -> None:
    assert to_search_query("Valter's deck", SearchScope.ANY) == '"\\"Valter\'s deck\\""'
