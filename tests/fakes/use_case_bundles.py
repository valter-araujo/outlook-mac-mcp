from datetime import datetime
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.clock import Clock
from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.list_upcoming_events import ListUpcomingEvents
from outlook_mac_mcp.application.ports.calendar_repository import CalendarRepository
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.application.preview_event import PreviewEvent
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.interface.mcp.calendar_write_use_cases import CalendarWriteUseCases
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


def calendar_write_use_cases(writer: CalendarWriter) -> UseCases:
    """A bundle with the write pair wired to `writer` over a fresh draft store."""
    drafts = DraftStore()
    write = CalendarWriteUseCases(
        preview_event=PreviewEvent(drafts), create_event=CreateEvent(drafts, writer)
    )
    bundle = _bundle(InMemoryMailRepository(), InMemoryCalendarRepository(), FixedClock(FROZEN_NOW))
    return UseCases(
        list_unread_emails=bundle.list_unread_emails,
        search_emails=bundle.search_emails,
        get_email=bundle.get_email,
        list_todays_events=bundle.list_todays_events,
        list_upcoming_events=bundle.list_upcoming_events,
        calendar_write=write,
    )


def _bundle(mail: MailRepository, calendar: CalendarRepository, clock: Clock) -> UseCases:
    return UseCases(
        list_unread_emails=ListUnreadEmails(mail),
        search_emails=SearchEmails(mail),
        get_email=GetEmail(mail),
        list_todays_events=ListTodaysEvents(calendar, clock),
        list_upcoming_events=ListUpcomingEvents(calendar, clock),
    )
