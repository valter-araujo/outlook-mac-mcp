from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email_filters import EmailFilters


class CountEmails:
    """How many emails match, and nothing else: no page, no items, one exact number."""

    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, filters: EmailFilters) -> int:
        return self._mail_repository.count_matching(filters)
