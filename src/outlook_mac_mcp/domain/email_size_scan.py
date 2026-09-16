from dataclasses import dataclass

from outlook_mac_mcp.domain.email_size import EmailSize


@dataclass(frozen=True, slots=True)
class EmailSizeScan:
    """What a walk over a folder yields: one EmailSize per message scanned, in scan order.

    `total` is how many messages matched in all; when the walk stopped at a ceiling the
    scan holds fewer than that and `coverage_is_complete` is false.
    """

    items: tuple[EmailSize, ...]
    total: int
    coverage_is_complete: bool

    @property
    def scanned(self) -> int:
        return len(self.items)
