from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.page import Page


class ListEmails:
    """The general listing: any combination of filters, sorted by date either way.

    This is the tool that can answer questions about a folder's extent, such as its
    oldest message, because unlike search its order is by date and its total is exact.
    """

    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: ListEmailsRequest) -> Page[Email]:
        return self._mail_repository.list_matching(request)
