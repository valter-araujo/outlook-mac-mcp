from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class Event:
    """One occurrence on the calendar.

    `start` and `end` are time zone-aware instants. An all-day event spans midnight to
    midnight in its own zone and is not treated differently here; `is_all_day` only tells
    a reader that the clock times are not meaningful.
    """

    id: str
    subject: str
    start: datetime
    end: datetime
    is_all_day: bool
    location: str
    organizer: EmailAddress
