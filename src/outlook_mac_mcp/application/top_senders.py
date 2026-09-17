from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.application.limits import ensure_limit_within_bounds
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.sender_ranking import SenderRanking, rank_senders

DEFAULT_TOP_SENDERS = 10
# Ten of the largest pages Graph serves: enough for most folders in one call, and a hard
# bound on what one tool call may cost when a folder holds far more.
SCAN_CEILING = 10_000


@dataclass(frozen=True, slots=True)
class TopSendersRequest:
    """The sender and attachment filters are absent on purpose: ranking senders by one
    sender is meaningless, and attachments say nothing about who writes most.
    """

    folders: tuple[str, ...] = (FolderName.INBOX,)
    is_read: bool | None = None
    received_after: datetime | None = None
    received_before: datetime | None = None
    limit: int = DEFAULT_TOP_SENDERS

    def __post_init__(self) -> None:
        ensure_limit_within_bounds(self.limit)
        self.filters()

    def filters(self) -> EmailFilters:
        return EmailFilters(
            folders=self.folders,
            is_read=self.is_read,
            received_after=self.received_after,
            received_before=self.received_before,
        )


class TopSenders:
    """Who writes most, computed here from a walk over the folder.

    The ceiling is a constructor argument so a test can drive the partial-coverage path
    without a ten-thousand-message fake.
    """

    def __init__(self, mail_repository: MailRepository, scan_ceiling: int = SCAN_CEILING) -> None:
        self._mail_repository = mail_repository
        self._scan_ceiling = scan_ceiling

    def execute(self, request: TopSendersRequest) -> SenderRanking:
        scan = self._mail_repository.scan_senders(request.filters(), self._scan_ceiling)
        return SenderRanking(
            senders=rank_senders(scan.senders, request.limit),
            scanned=scan.scanned,
            total=scan.total,
            coverage_is_complete=scan.coverage_is_complete,
        )
