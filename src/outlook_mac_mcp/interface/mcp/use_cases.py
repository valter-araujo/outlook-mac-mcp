from dataclasses import dataclass

from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.list_upcoming_events import ListUpcomingEvents
from outlook_mac_mcp.application.search_emails import SearchEmails


@dataclass(frozen=True, slots=True)
class UseCases:
    """Everything the server registers, handed over as one object.

    Defined here rather than in the composition root because the interface layer is
    what consumes it; the root only fills it in.
    """

    list_unread_emails: ListUnreadEmails
    search_emails: SearchEmails
    get_email: GetEmail
    list_todays_events: ListTodaysEvents
    list_upcoming_events: ListUpcomingEvents
