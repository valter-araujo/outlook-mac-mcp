from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.list_upcoming_events import ListUpcomingEvents
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.calendar_repository import GraphCalendarRepository
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.mail_repository import GraphMailRepository
from outlook_mac_mcp.infrastructure.settings import Settings
from outlook_mac_mcp.infrastructure.system_clock import SystemClock
from outlook_mac_mcp.interface.mcp.use_cases import UseCases


def build_use_cases(settings: Settings) -> UseCases:
    """Wire the use cases to the Graph adapters.

    The composition root sits outside the layers because it is the one place allowed to
    know every one of them; nothing here decides behaviour, it only connects. One client
    serves both repositories, so one token and one HTTP client serve the process.
    """
    client = GraphClient(DeviceCodeAuthenticator.from_settings(settings))
    mail = GraphMailRepository(client)
    calendar = GraphCalendarRepository(client, settings.timezone)
    clock = SystemClock(settings.timezone)
    return UseCases(
        list_unread_emails=ListUnreadEmails(mail),
        search_emails=SearchEmails(mail),
        get_email=GetEmail(mail),
        list_todays_events=ListTodaysEvents(calendar, clock),
        list_upcoming_events=ListUpcomingEvents(calendar, clock),
    )
