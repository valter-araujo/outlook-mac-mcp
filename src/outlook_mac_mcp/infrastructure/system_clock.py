from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo


@dataclass(frozen=True, slots=True)
class SystemClock:
    """The real clock, read in the resolved zone; the only place datetime.now is called."""

    timezone: ZoneInfo

    def now(self) -> datetime:
        return datetime.now(self.timezone)
