from dataclasses import dataclass

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError

MIN_SUBJECT_LENGTH = 1
MAX_SUBJECT_LENGTH = 255
MAX_BODY_LENGTH = 32_768
MAX_RECIPIENTS = 50


@dataclass(frozen=True, slots=True)
class NewMailDraft:
    """A message as the user wants it saved, validated before anything is previewed.

    A draft is only ever saved, never sent: this project has no send path, and the
    person presses send in Outlook themselves. That is the defence against a message
    body that talks a model into mailing someone — the worst a hostile instruction can
    produce is a draft the user sees before it goes anywhere.

    `body` is plain text only, for the same reason NewEvent's is: it can originate from
    an email or a web page, and accepting markup would reopen the injection surface
    get_email's body warning already covers.
    """

    subject: str
    body: str
    to: tuple[EmailAddress, ...]
    cc: tuple[EmailAddress, ...] = ()

    def __post_init__(self) -> None:
        if not MIN_SUBJECT_LENGTH <= len(self.subject) <= MAX_SUBJECT_LENGTH:
            raise InvalidRequestError(
                f"subject must be between {MIN_SUBJECT_LENGTH} and {MAX_SUBJECT_LENGTH} characters"
            )
        if len(self.body) > MAX_BODY_LENGTH:
            raise InvalidRequestError(f"body must be at most {MAX_BODY_LENGTH} characters")
        if not self.to:
            raise InvalidRequestError("a draft needs at least one recipient in to")
        if len(self.to) + len(self.cc) > MAX_RECIPIENTS:
            raise InvalidRequestError(f"at most {MAX_RECIPIENTS} recipients are allowed")
