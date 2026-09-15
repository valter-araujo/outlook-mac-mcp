from datetime import datetime
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.clock import Clock
from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.list_upcoming_events import ListUpcomingEvents
from outlook_mac_mcp.application.ports.calendar_repository import CalendarRepository
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.interface.mcp.use_cases import UseCases
from tests.fakes.fixed_clock import FixedClock
from tests.fakes.in_memory_calendar_repository import InMemoryCalendarRepository
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

FROZEN_NOW = datetime(2026, 9, 15, 15, 42, tzinfo=ZoneInfo("America/Sao_Paulo"))


def mail_only_use_cases(repository: MailRepository) -> UseCases:
    """A bundle for tests of the mail tools; the calendar side is empty and frozen."""
    return _bundle(repository, InMemoryCalendarRepository(), FixedClock(FROZEN_NOW))


def calendar_only_use_cases(repository: CalendarRepository, clock: Clock) -> UseCases:
    """A bundle for tests of the calendar tools; the mail side is empty."""
    return _bundle(InMemoryMailRepository(), repository, clock)


def _bundle(mail: MailRepository, calendar: CalendarRepository, clock: Clock) -> UseCases:
    return UseCases(
        list_unread_emails=ListUnreadEmails(mail),
        search_emails=SearchEmails(mail),
        get_email=GetEmail(mail),
        list_todays_events=ListTodaysEvents(calendar, clock),
        list_upcoming_events=ListUpcomingEvents(calendar, clock),
    )
