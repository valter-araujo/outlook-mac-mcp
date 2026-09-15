from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True, slots=True)
class FixedClock:
    """A clock stopped at one instant, so a test decides what today is."""

    instant: datetime

    def now(self) -> datetime:
        return self.instant
