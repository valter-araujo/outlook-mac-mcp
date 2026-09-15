from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.clock import Clock
from outlook_mac_mcp.infrastructure.system_clock import SystemClock

LISBON = ZoneInfo("Europe/Lisbon")


def test_satisfies_the_clock_port() -> None:
    clock: Clock = SystemClock(LISBON)

    assert clock.now().tzinfo == LISBON


def test_reads_the_current_instant_in_the_given_zone() -> None:
    before = datetime.now(LISBON)

    now = SystemClock(LISBON).now()

    assert now.tzinfo == LISBON
    assert before <= now <= before + timedelta(seconds=5)
