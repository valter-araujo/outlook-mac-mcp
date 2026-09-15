from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.sender_count import SenderCount
from outlook_mac_mcp.domain.sender_ranking import rank_senders


def address(value: str, name: str = "") -> EmailAddress:
    return EmailAddress(address=value, display_name=name)


def test_ranks_by_count_most_first() -> None:
    senders = [address("b@x.io"), address("a@x.io"), address("b@x.io"), address("c@x.io")]

    ranked = rank_senders(senders, limit=10)

    assert [(item.sender.address, item.count) for item in ranked] == [
        ("b@x.io", 2),
        ("a@x.io", 1),
        ("c@x.io", 1),
    ]


def test_breaks_ties_by_address_so_the_order_is_stable() -> None:
    senders = [address("zed@x.io"), address("amy@x.io"), address("mid@x.io")]

    ranked = rank_senders(senders, limit=10)

    assert [item.sender.address for item in ranked] == ["amy@x.io", "mid@x.io", "zed@x.io"]


def test_merges_spellings_that_differ_only_in_case_keeping_the_first() -> None:
    senders = [address("Ana@Example.com"), address("ana@example.com"), address("ANA@EXAMPLE.COM")]

    ranked = rank_senders(senders, limit=10)

    assert ranked == (SenderCount(sender=address("Ana@Example.com"), count=3),)


def test_keeps_the_first_non_empty_display_name() -> None:
    senders = [address("ana@x.io"), address("ana@x.io", "Ana Lima"), address("ana@x.io", "A. Lima")]

    ranked = rank_senders(senders, limit=10)

    assert ranked[0].sender == address("ana@x.io", "Ana Lima")


def test_scans_past_a_message_with_no_sender() -> None:
    ranked = rank_senders([address(""), address("ana@x.io"), address("")], limit=10)

    assert [item.sender.address for item in ranked] == ["ana@x.io"]


def test_cuts_the_ranking_at_the_limit() -> None:
    senders = [address(f"{letter}@x.io") for letter in "abcde"]

    assert len(rank_senders(senders, limit=2)) == 2


def test_ranks_nobody_when_nothing_was_scanned() -> None:
    assert rank_senders([], limit=10) == ()
