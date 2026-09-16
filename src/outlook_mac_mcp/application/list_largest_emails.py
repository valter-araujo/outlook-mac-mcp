from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.application.limits import ensure_limit_within_bounds
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_size_ranking import EmailSizeRanking, rank_by_size
from outlook_mac_mcp.domain.folder_name import FolderName

DEFAULT_LARGEST_EMAILS = 10
# Same bound and page-size pattern as the sender scan: ten of the largest pages Graph
# serves, enough for most folders in one call, and a hard cap on what one call can cost.
SCAN_CEILING = 10_000


@dataclass(frozen=True, slots=True)
class ListLargestEmailsRequest:
    """No sender or attachment filter: a size ranking is about the folder, not one
    correspondent, and whether an email has an attachment is already implied by its size.
    """

    folder: FolderName = FolderName.INBOX
    is_read: bool | None = None
    received_after: datetime | None = None
    limit: int = DEFAULT_LARGEST_EMAILS

    def __post_init__(self) -> None:
        ensure_limit_within_bounds(self.limit)
        self.filters()

    def filters(self) -> EmailFilters:
        return EmailFilters(
            folder=self.folder, is_read=self.is_read, received_after=self.received_after
        )


class ListLargestEmails:
    """The largest emails in a folder, computed here from a walk over the folder.

    Graph has no $orderby for an extended property, so the size cannot be sorted on the
    wire; the walk collects sizes and the domain sorts them. The ceiling is a constructor
    argument so a test can drive the partial-coverage path without a huge fake.
    """

    def __init__(self, mail_repository: MailRepository, scan_ceiling: int = SCAN_CEILING) -> None:
        self._mail_repository = mail_repository
        self._scan_ceiling = scan_ceiling

    def execute(self, request: ListLargestEmailsRequest) -> EmailSizeRanking:
        scan = self._mail_repository.scan_email_sizes(request.filters(), self._scan_ceiling)
        return EmailSizeRanking(
            items=rank_by_size(scan.items, request.limit),
            scanned=scan.scanned,
            skipped=scan.skipped,
            total=scan.total,
            coverage_is_complete=scan.coverage_is_complete,
        )
