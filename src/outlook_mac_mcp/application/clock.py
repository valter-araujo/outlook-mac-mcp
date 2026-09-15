from datetime import datetime
from typing import Protocol


class Clock(Protocol):
    """Where a use case learns what time it is.

    Injected so a test can pin the moment, and so "today" is decided once, from a single
    reading, rather than by scattered calls that could straddle midnight.
    """

    def now(self) -> datetime:
        """The current instant, time zone-aware, in the zone the user reads the calendar in."""
        ...
