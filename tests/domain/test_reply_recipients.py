from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import (
    ReplyRecipients,
    address_key,
    resolve_reply_recipients,
)
from outlook_mac_mcp.domain.reply_source import ReplySource

ME = "me@example.com"
OWN = frozenset({ME})


def addresses(*values: str) -> tuple[EmailAddress, ...]:
    return tuple(EmailAddress(address=value) for value in values)


def source(
    sender: str = "ana@example.com",
    to: tuple[str, ...] = (ME,),
    cc: tuple[str, ...] = (),
    reply_to: tuple[str, ...] = (),
) -> ReplySource:
    return ReplySource(
        id="AAMk1",
        subject="Interview",
        sender=EmailAddress(address=sender),
        reply_to=addresses(*reply_to),
        to=addresses(*to),
        cc=addresses(*cc),
    )


def to_line(recipients: ReplyRecipients) -> list[str]:
    return [recipient.address for recipient in recipients.to]


def cc_line(recipients: ReplyRecipients) -> list[str]:
    return [recipient.address for recipient in recipients.cc]


def test_reply_goes_to_the_sender_only() -> None:
    message = source(to=(ME, "bo@example.com"), cc=("cy@example.com",))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY, OWN)

    assert (to_line(recipients), cc_line(recipients)) == (["ana@example.com"], [])


def test_reply_all_keeps_every_original_recipient_except_the_user() -> None:
    message = source(to=(ME, "bo@example.com"), cc=("cy@example.com", "di@example.com"))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY_ALL, OWN)

    assert to_line(recipients) == ["ana@example.com", "bo@example.com"]
    assert cc_line(recipients) == ["cy@example.com", "di@example.com"]


def test_reply_all_removes_the_user_from_cc_too() -> None:
    message = source(to=("bo@example.com",), cc=(ME, "cy@example.com"))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY_ALL, OWN)

    assert cc_line(recipients) == ["cy@example.com"]


def test_excludes_the_user_regardless_of_letter_case() -> None:
    message = source(to=("Me@Example.COM", "bo@example.com"))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY_ALL, OWN)

    assert to_line(recipients) == ["ana@example.com", "bo@example.com"]


def test_lists_an_address_once_when_it_appears_on_several_lines() -> None:
    message = source(to=("ana@example.com", "bo@example.com"), cc=("BO@example.com",))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY_ALL, OWN)

    assert (to_line(recipients), cc_line(recipients)) == (
        ["ana@example.com", "bo@example.com"],
        [],
    )


def test_reply_to_replaces_the_sender_as_the_addressee() -> None:
    message = source(sender="robot@example.com", reply_to=("recruiter@example.com",))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY, OWN)

    assert to_line(recipients) == ["recruiter@example.com"]


def test_reply_to_the_users_own_message_leaves_no_addressee() -> None:
    message = source(sender=ME, to=("bo@example.com",))

    recipients = resolve_reply_recipients(message, ReplyMode.REPLY, OWN)

    assert recipients.to == ()


def test_extras_are_appended_once() -> None:
    resolved = resolve_reply_recipients(source(), ReplyMode.REPLY, OWN)

    recipients = resolved.with_extra(
        addresses("ANA@example.com", "eve@example.com"), addresses("eve@example.com")
    )

    assert (to_line(recipients), cc_line(recipients)) == (
        ["ana@example.com", "eve@example.com"],
        [],
    )


def test_an_extra_naming_the_user_is_kept_because_it_was_asked_for() -> None:
    resolved = resolve_reply_recipients(source(), ReplyMode.REPLY, OWN)

    recipients = resolved.with_extra((), addresses(ME))

    assert cc_line(recipients) == [ME]


def test_address_key_ignores_case_and_surrounding_spaces() -> None:
    assert address_key("  Ana@Example.com ") == "ana@example.com"
