import pytest

from outlook_mac_mcp.infrastructure.graph.search_query import to_literal_phrase


def test_wraps_a_plain_term_in_quotes() -> None:
    assert to_literal_phrase("quarterly review") == '"quarterly review"'


@pytest.mark.parametrize(
    "operator_like",
    [
        "from:ceo@example.com",
        "subject:payroll",
        "AND",
        "OR",
        "NOT",
        "deck AND from:ana@example.com",
        "received>=2026-01-01",
        "(a OR b)",
    ],
)
def test_keeps_text_that_reads_like_a_query_inside_the_phrase(operator_like: str) -> None:
    assert to_literal_phrase(operator_like) == f'"{operator_like}"'


def test_escapes_an_embedded_quote_so_the_phrase_cannot_be_closed() -> None:
    assert to_literal_phrase('say "hello"') == '"say \\"hello\\""'


def test_escapes_a_breakout_attempt_that_appends_an_operator() -> None:
    breakout = 'deck" AND from:ceo@example.com "'

    phrase = to_literal_phrase(breakout)

    assert phrase == '"deck\\" AND from:ceo@example.com \\""'
    assert phrase.startswith('"')
    assert phrase.endswith('"')


def test_escapes_backslashes_before_quotes() -> None:
    """A term ending in a backslash must not consume the escape and close the phrase."""
    assert to_literal_phrase("path\\") == '"path\\\\"'


def test_a_trailing_backslash_cannot_smuggle_an_operator() -> None:
    phrase = to_literal_phrase('a\\" AND from:ceo@example.com')

    assert phrase == '"a\\\\\\" AND from:ceo@example.com"'


def test_leaves_a_lone_colon_alone() -> None:
    assert to_literal_phrase("ratio 3:1") == '"ratio 3:1"'


def test_keeps_an_empty_looking_term_quoted() -> None:
    assert to_literal_phrase(" ") == '" "'
