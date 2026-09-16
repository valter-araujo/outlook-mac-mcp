from collections.abc import Iterable
from dataclasses import dataclass

from outlook_mac_mcp.domain.email_size import EmailSize


@dataclass(frozen=True, slots=True)
class EmailSizeRanking:
    items: tuple[EmailSize, ...]
    scanned: int
    skipped: int
    total: int
    coverage_is_complete: bool


def rank_by_size(items: Iterable[EmailSize], limit: int) -> tuple[EmailSize, ...]:
    """Largest first. Graph has no $orderby for an extended property, so this is the
    sort; ties keep scan order, since Python's sort is stable.
    """
    return tuple(sorted(items, key=lambda item: item.size_bytes, reverse=True)[:limit])
