from collections.abc import Iterable
from dataclasses import dataclass

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_source import ReplySource


def address_key(address: str) -> str:
    """Email addresses compare case-insensitively in practice, so every exclusion and
    deduplication goes through this one key.
    """
    return address.strip().casefold()


@dataclass(frozen=True, slots=True)
class ReplyRecipients:
    """The final To and Cc lines of a reply, each address appearing once across both."""

    to: tuple[EmailAddress, ...]
    cc: tuple[EmailAddress, ...]

    def with_extra(
        self, extra_to: tuple[EmailAddress, ...], extra_cc: tuple[EmailAddress, ...]
    ) -> "ReplyRecipients":
        """Extras the user names are kept even when one is their own address: an
        explicit request wins over the self-exclusion rule.
        """
        to = _unique((*self.to, *extra_to), frozenset())
        cc = _unique((*self.cc, *extra_cc), _keys(to))
        return ReplyRecipients(to=to, cc=cc)


def resolve_reply_recipients(
    source: ReplySource, mode: ReplyMode, own_addresses: frozenset[str]
) -> ReplyRecipients:
    """Who a reply goes to, computed here rather than left to Graph so the preview shows
    exactly the list that will be saved.

    The addressee is Reply-To when the message has one, otherwise the sender. Reply all
    adds every original To recipient to the To line and keeps the original Cc line as Cc.
    The user's own addresses (`own_addresses`, already keyed with address_key) are
    removed throughout, and an address already on To never repeats on Cc.
    """
    addressees = source.reply_to or (source.sender,)
    if mode is ReplyMode.REPLY:
        return ReplyRecipients(to=_unique(addressees, own_addresses), cc=())
    to = _unique((*addressees, *source.to), own_addresses)
    cc = _unique(source.cc, own_addresses | _keys(to))
    return ReplyRecipients(to=to, cc=cc)


def _unique(
    addresses: Iterable[EmailAddress], excluded: frozenset[str]
) -> tuple[EmailAddress, ...]:
    seen = set(excluded)
    kept: list[EmailAddress] = []
    for recipient in addresses:
        key = address_key(recipient.address)
        if not key or key in seen:
            continue
        seen.add(key)
        kept.append(recipient)
    return tuple(kept)


def _keys(addresses: tuple[EmailAddress, ...]) -> frozenset[str]:
    return frozenset(address_key(recipient.address) for recipient in addresses)
