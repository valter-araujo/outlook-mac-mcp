from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.errors import InvalidRequestError


@dataclass(frozen=True, slots=True)
class TimeWindow:
    """A half-open range of instants: `start` is inside, `end` is not.

    Both ends must carry a time zone. A naive datetime has no instant to compare with,
    and the calendar backend needs an offset to interpret the range at all.
    """

    start: datetime
    end: datetime

    def __post_init__(self) -> None:
        if self.start.tzinfo is None or self.end.tzinfo is None:
            raise InvalidRequestError("a time window needs time zone-aware boundaries")
        if not self.start < self.end:
            raise InvalidRequestError("a time window must start before it ends")

    def overlaps(self, start: datetime, end: datetime) -> bool:
        """Whether an occurrence from `start` to `end` shares any instant with the window.

        Half-open on both sides, so an event ending exactly at the window's start, or
        starting exactly at its end, is outside: a midnight-to-midnight all-day event
        belongs to one day, never to two.
        """
        return start < self.end and end > self.start
