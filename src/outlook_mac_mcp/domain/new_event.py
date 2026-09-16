from dataclasses import dataclass
from datetime import datetime, time

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError

MIN_SUBJECT_LENGTH = 1
MAX_SUBJECT_LENGTH = 255
MAX_ATTENDEES = 50
MAX_BODY_LENGTH = 32_768


@dataclass(frozen=True, slots=True)
class NewEvent:
    """An event as the user wants it, validated before anything is drafted or sent.

    Rules live here and not in the tool's input model so a draft can only ever hold an
    event the calendar would accept, whichever interface produced it.

    `body` is plain text only: there is no field or code path anywhere in this project
    that lets a caller choose HTML instead. A body can originate from an email or a web
    page — the same untrusted-content surface get_email's body warning already covers —
    and accepting markup there would be the same injection risk in a different tool.
    """

    subject: str
    start: datetime
    end: datetime
    is_all_day: bool = False
    location: str = ""
    body: str = ""
    attendees: tuple[EmailAddress, ...] = ()

    def __post_init__(self) -> None:
        if not MIN_SUBJECT_LENGTH <= len(self.subject) <= MAX_SUBJECT_LENGTH:
            raise InvalidRequestError(
                f"subject must be between {MIN_SUBJECT_LENGTH} and {MAX_SUBJECT_LENGTH} characters"
            )
        if self.start.tzinfo is None or self.end.tzinfo is None:
            raise InvalidRequestError("start and end must carry a time zone")
        if not self.start < self.end:
            raise InvalidRequestError("start must be before end")
        if len(self.attendees) > MAX_ATTENDEES:
            raise InvalidRequestError(f"at most {MAX_ATTENDEES} attendees are allowed")
        if len(self.body) > MAX_BODY_LENGTH:
            raise InvalidRequestError(f"body must be at most {MAX_BODY_LENGTH} characters")
        if self.is_all_day:
            _ensure_midnight_boundaries(self.start, self.end)


def _ensure_midnight_boundaries(start: datetime, end: datetime) -> None:
    """The calendar backend refuses an all-day event that does not run midnight to midnight,
    so the rule is enforced here rather than discovered as a remote error.
    """
    if (
        start.timetz().replace(tzinfo=None) != time.min
        or end.timetz().replace(tzinfo=None) != time.min
    ):
        raise InvalidRequestError("an all-day event must start and end at midnight")
