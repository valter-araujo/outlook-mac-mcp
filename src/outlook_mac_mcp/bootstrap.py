from outlook_mac_mcp.application.count_emails import CountEmails
from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.delete_event import DeleteEvent
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_deletion_draft import EventDeletionDraft
from outlook_mac_mcp.application.event_draft import EventDraft
from outlook_mac_mcp.application.event_update_draft import EventUpdateDraft
from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_custom_folders import ListCustomFolders
from outlook_mac_mcp.application.list_emails import ListEmails
from outlook_mac_mcp.application.list_folders import ListFolders
from outlook_mac_mcp.application.list_largest_emails import ListLargestEmails
from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.list_upcoming_events import ListUpcomingEvents
from outlook_mac_mcp.application.preview_event import PreviewEvent
from outlook_mac_mcp.application.preview_event_deletion import PreviewEventDeletion
from outlook_mac_mcp.application.preview_event_update import PreviewEventUpdate
from outlook_mac_mcp.application.search_contacts import SearchContacts
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.application.top_senders import TopSenders
from outlook_mac_mcp.application.update_event import UpdateEvent
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.calendar_repository import GraphCalendarRepository
from outlook_mac_mcp.infrastructure.graph.calendar_writer import GraphCalendarWriter
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.contact_repository import GraphContactRepository
from outlook_mac_mcp.infrastructure.graph.mail_folder_repository import GraphMailFolderRepository
from outlook_mac_mcp.infrastructure.graph.mail_repository import GraphMailRepository
from outlook_mac_mcp.infrastructure.settings import Settings
from outlook_mac_mcp.infrastructure.system_clock import SystemClock
from outlook_mac_mcp.interface.mcp.calendar_write_use_cases import CalendarWriteUseCases
from outlook_mac_mcp.interface.mcp.use_cases import UseCases


def build_use_cases(settings: Settings) -> UseCases:
    """Wire the use cases to the Graph adapters.

    The composition root sits outside the layers because it is the one place allowed to
    know every one of them; nothing here decides behaviour, it only connects. One client
    serves both repositories, so one token and one HTTP client serve the process.
    """
    client = GraphClient(DeviceCodeAuthenticator.from_settings(settings))
    mail = GraphMailRepository(client)
    folders = GraphMailFolderRepository(client)
    contacts = GraphContactRepository(client)
    calendar = GraphCalendarRepository(client, settings.timezone)
    clock = SystemClock(settings.timezone)
    write = _calendar_write(client, settings) if settings.calendar_write_enabled else None
    return UseCases(
        list_unread_emails=ListUnreadEmails(mail),
        search_emails=SearchEmails(mail),
        get_email=GetEmail(mail),
        list_todays_events=ListTodaysEvents(calendar, clock),
        list_upcoming_events=ListUpcomingEvents(calendar, clock),
        list_emails=ListEmails(mail),
        count_emails=CountEmails(mail),
        top_senders=TopSenders(mail),
        list_largest_emails=ListLargestEmails(mail),
        list_folders=ListFolders(folders),
        list_custom_folders=ListCustomFolders(folders),
        search_contacts=SearchContacts(contacts),
        calendar_write=write,
    )


def _calendar_write(client: GraphClient, settings: Settings) -> CalendarWriteUseCases:
    """One draft store per pair, each shared only within that pair, is what makes a
    token single-use and non-transferable between create, update and deletion.
    """
    create_drafts: DraftStore[EventDraft] = DraftStore()
    update_drafts: DraftStore[EventUpdateDraft] = DraftStore()
    writer = GraphCalendarWriter(client, settings.timezone)
    preview_event_deletion, delete_event = _calendar_deletion(writer, settings)
    return CalendarWriteUseCases(
        preview_event=PreviewEvent(create_drafts),
        create_event=CreateEvent(create_drafts, writer),
        preview_event_update=PreviewEventUpdate(writer, update_drafts),
        update_event=UpdateEvent(update_drafts, writer),
        preview_event_deletion=preview_event_deletion,
        delete_event=delete_event,
    )


def _calendar_deletion(
    writer: GraphCalendarWriter, settings: Settings
) -> tuple[PreviewEventDeletion, DeleteEvent] | tuple[None, None]:
    """Gated by its own flag on top of calendar write, so the one irreversible action
    is never enabled just because write is.
    """
    if not settings.calendar_delete_enabled:
        return None, None
    deletion_drafts: DraftStore[EventDeletionDraft] = DraftStore()
    return PreviewEventDeletion(writer, deletion_drafts), DeleteEvent(deletion_drafts, writer)
