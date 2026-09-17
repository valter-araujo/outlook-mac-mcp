from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email_search_page import EmailSearchPage


class SearchEmails:
    """Find emails in one folder whose text matches a term.

    Results come back in the backend's relevance order, not newest first: Graph cannot
    combine `$search` with `$orderby`, and sorting a single relevance-ranked page by date
    here would reorder an arbitrary subset and read as if it were the newest mail.
    """

    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: SearchEmailsRequest) -> EmailSearchPage:
        return self._mail_repository.search(request)
