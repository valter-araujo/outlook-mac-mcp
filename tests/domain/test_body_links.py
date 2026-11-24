import pytest

from outlook_mac_mcp.domain.body_links import find_links


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("no links here", ()),
        ("see https://example.com/x?d=secret", ("https://example.com/x?d=secret",)),
        ("plain http://example.com too", ("http://example.com",)),
        ("go to www.example.com.", ("www.example.com",)),
        ("(https://example.com/a)", ("https://example.com/a",)),
        ("HTTPS://EXAMPLE.COM/A", ("HTTPS://EXAMPLE.COM/A",)),
        ('<a href="https://example.com/p">x</a>', ("https://example.com/p",)),
    ],
)
def test_finds_web_addresses(text: str, expected: tuple[str, ...]) -> None:
    assert find_links(text) == expected


def test_lists_each_link_once_in_order() -> None:
    text = "https://b.example.com then https://a.example.com and https://b.example.com"

    assert find_links(text) == ("https://b.example.com", "https://a.example.com")


def test_ignores_a_bare_scheme() -> None:
    assert find_links("type https:// and then the host") == ()
