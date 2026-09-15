from dataclasses import dataclass

from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.errors import InvalidRequestError

# Graph message ids are opaque and their length varies with the mailbox and with
# immutable-id mode, so the bound exists to keep unbounded input out of a URL path,
# not to describe a format.
MAX_EMAIL_ID_LENGTH = 1024


@dataclass(frozen=True, slots=True)
class GetEmailRequest:
    email_id: str

    def __post_init__(self) -> None:
        if not self.email_id:
            raise InvalidRequestError("email_id must not be empty")
        if len(self.email_id) > MAX_EMAIL_ID_LENGTH:
            raise InvalidRequestError(f"email_id must be at most {MAX_EMAIL_ID_LENGTH} characters")


class GetEmail:
    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: GetEmailRequest) -> EmailDetail:
        return self._mail_repository.get_by_id(request.email_id)
