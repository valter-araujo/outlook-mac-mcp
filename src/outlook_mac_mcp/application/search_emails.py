from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName

MIN_TERM_LENGTH = 1
MAX_TERM_LENGTH = 200


@dataclass(frozen=True, slots=True)
class SearchEmailsRequest:
    term: str
    folder: FolderName = FolderName.INBOX
    limit: int = DEFAULT_LIMIT

    def __post_init__(self) -> None:
        if not MIN_TERM_LENGTH <= len(self.term) <= MAX_TERM_LENGTH:
            raise InvalidRequestError(
                f"term must be between {MIN_TERM_LENGTH} and {MAX_TERM_LENGTH} characters"
            )
        ensure_limit_within_bounds(self.limit)


class SearchEmails:
    """Find emails in one folder whose text matches a term.

    Results come back in the backend's relevance order, not newest first: Graph cannot
    combine `$search` with `$orderby`, and sorting a single relevance-ranked page by date
    here would reorder an arbitrary subset and read as if it were the newest mail.
    """

    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: SearchEmailsRequest) -> tuple[Email, ...]:
        return self._mail_repository.search(request.folder, request.term, request.limit)
