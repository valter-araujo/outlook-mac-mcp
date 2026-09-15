from collections.abc import Iterable
from dataclasses import dataclass

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.sender_count import SenderCount


@dataclass(frozen=True, slots=True)
class SenderRanking:
    senders: tuple[SenderCount, ...]
    scanned: int
    total: int
    coverage_is_complete: bool


def rank_senders(senders: Iterable[EmailAddress], limit: int) -> tuple[SenderCount, ...]:
    """Count each address, most emails first; ties break by address so the order is stable.

    Addresses are merged without regard to case, since mail systems treat them that way,
    and the first spelling and the first non-empty display name seen are the ones kept.
    A message with no sender is scanned but ranks nobody.
    """
    counts: dict[str, int] = {}
    first_seen: dict[str, EmailAddress] = {}
    for sender in senders:
        if not sender.address:
            continue
        key = sender.address.casefold()
        counts[key] = counts.get(key, 0) + 1
        first_seen[key] = _prefer_named(first_seen.get(key), sender)
    ranked = sorted(counts, key=lambda key: (-counts[key], key))
    return tuple(SenderCount(sender=first_seen[key], count=counts[key]) for key in ranked[:limit])


def _prefer_named(kept: EmailAddress | None, seen: EmailAddress) -> EmailAddress:
    if kept is None:
        return seen
    if kept.display_name or not seen.display_name:
        return kept
    return EmailAddress(address=kept.address, display_name=seen.display_name)
