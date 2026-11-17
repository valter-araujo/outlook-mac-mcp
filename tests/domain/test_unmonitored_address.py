import pytest

from outlook_mac_mcp.domain.unmonitored_address import looks_unmonitored


@pytest.mark.parametrize(
    "address",
    [
        "noreply@example.com",
        "no-reply@example.com",
        "do_not_reply@example.com",
        "DoNotReply@example.com",
        "mailer-daemon@example.com",
        "notifications@example.com",
        "jobs-noreply@example.com",
    ],
)
def test_flags_a_no_reply_local_part(address: str) -> None:
    assert looks_unmonitored(address) is True


@pytest.mark.parametrize(
    "address",
    ["recruiting@email.example.com", "talent@mail3.example.com", "jobs@em.example.co.uk"],
)
def test_flags_an_address_on_a_bulk_sending_host(address: str) -> None:
    assert looks_unmonitored(address) is True


@pytest.mark.parametrize(
    "address",
    ["ana@example.com", "ana@mail.com", "ana@team.example.com", "eleanor@example.com"],
)
def test_does_not_flag_an_ordinary_address(address: str) -> None:
    assert looks_unmonitored(address) is False


def test_does_not_flag_something_that_is_not_an_address() -> None:
    assert looks_unmonitored("not an address") is False
