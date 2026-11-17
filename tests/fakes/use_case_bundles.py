from dataclasses import replace
from datetime import datetime
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.clock import Clock
from outlook_mac_mcp.application.count_emails import CountEmails
from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.create_mail_draft import CreateMailDraft
from outlook_mac_mcp.application.create_reply_draft import CreateReplyDraft
from outlook_mac_mcp.application.delete_draft import DeleteDraft
from outlook_mac_mcp.application.delete_event import DeleteEvent
from outlook_mac_mcp.application.draft_deletion_preview import DraftDeletionPreview
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
from outlook_mac_mcp.application.mail_draft_preview import MailDraftPreview
from outlook_mac_mcp.application.ports.calendar_repository import CalendarRepository
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.application.ports.contact_repository import ContactRepository
from outlook_mac_mcp.application.ports.draft_deleter import DraftDeleter
from outlook_mac_mcp.application.ports.mail_draft_writer import MailDraftWriter
from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.application.ports.reply_draft_writer import ReplyDraftWriter
from outlook_mac_mcp.application.preview_draft_deletion import PreviewDraftDeletion
from outlook_mac_mcp.application.preview_event import PreviewEvent
from outlook_mac_mcp.application.preview_event_deletion import PreviewEventDeletion
from outlook_mac_mcp.application.preview_event_update import PreviewEventUpdate
from outlook_mac_mcp.application.preview_mail_draft import PreviewMailDraft
from outlook_mac_mcp.application.preview_reply_draft import PreviewReplyDraft
from outlook_mac_mcp.application.reply_draft_preview import ReplyDraftPreview
from outlook_mac_mcp.application.resolve_folders import ResolveFolders
from outlook_mac_mcp.application.search_contacts import SearchContacts
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.application.top_senders import TopSenders
from outlook_mac_mcp.application.update_event import UpdateEvent
from outlook_mac_mcp.interface.mcp.calendar_write_use_cases import CalendarWriteUseCases
from outlook_mac_mcp.interface.mcp.mail_draft_use_cases import MailDraftUseCases
from outlook_mac_mcp.interface.mcp.use_cases import UseCases
from tests.fakes.fixed_clock import FixedClock
from tests.fakes.in_memory_calendar_repository import InMemoryCalendarRepository
from tests.fakes.in_memory_contact_repository import InMemoryContactRepository
from tests.fakes.in_memory_mail_folder_repository import InMemoryMailFolderRepository
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.in_memory_reply_draft_writer import InMemoryReplyDraftWriter

FROZEN_NOW = datetime(2026, 9, 15, 15, 42, tzinfo=ZoneInfo("America/Sao_Paulo"))


def mail_only_use_cases(
    repository: MailRepository, folder_repository: MailFolderRepository | None = None
) -> UseCases:
    """A bundle for tests of the mail tools; the calendar side is empty and frozen.

    `folder_repository` backs resolve_folders -- pass one pre-populated with custom
    folders (InMemoryMailFolderRepository.set_custom) to test a custom folder path
    resolving through a mail tool; omitted, custom paths never match anything.
    """
    return _bundle(
        repository, InMemoryCalendarRepository(), FixedClock(FROZEN_NOW), folder_repository
    )


def calendar_only_use_cases(repository: CalendarRepository, clock: Clock) -> UseCases:
    """A bundle for tests of the calendar tools; the mail side is empty."""
    return _bundle(InMemoryMailRepository(), repository, clock)


def calendar_write_use_cases(writer: CalendarWriter) -> UseCases:
    """A bundle with all three write pairs wired to `writer`, each over its own fresh
    draft store.
    """
    create_drafts: DraftStore[EventDraft] = DraftStore()
    update_drafts: DraftStore[EventUpdateDraft] = DraftStore()
    deletion_drafts: DraftStore[EventDeletionDraft] = DraftStore()
    write = CalendarWriteUseCases(
        preview_event=PreviewEvent(create_drafts),
        create_event=CreateEvent(create_drafts, writer),
        preview_event_update=PreviewEventUpdate(writer, update_drafts),
        update_event=UpdateEvent(update_drafts, writer),
        preview_event_deletion=PreviewEventDeletion(writer, deletion_drafts),
        delete_event=DeleteEvent(deletion_drafts, writer),
    )
    bundle = _bundle(InMemoryMailRepository(), InMemoryCalendarRepository(), FixedClock(FROZEN_NOW))
    return UseCases(
        list_unread_emails=bundle.list_unread_emails,
        search_emails=bundle.search_emails,
        get_email=bundle.get_email,
        list_todays_events=bundle.list_todays_events,
        list_upcoming_events=bundle.list_upcoming_events,
        list_emails=bundle.list_emails,
        count_emails=bundle.count_emails,
        top_senders=bundle.top_senders,
        list_largest_emails=bundle.list_largest_emails,
        list_folders=bundle.list_folders,
        list_custom_folders=bundle.list_custom_folders,
        resolve_folders=bundle.resolve_folders,
        search_contacts=bundle.search_contacts,
        calendar_write=write,
    )


def calendar_write_use_cases_without_deletion(writer: CalendarWriter) -> UseCases:
    """A bundle with create and update wired to `writer`, deletion left unset — as if
    OUTLOOK_MCP_ENABLE_CALENDAR_DELETE were off while write is on.
    """
    create_drafts: DraftStore[EventDraft] = DraftStore()
    update_drafts: DraftStore[EventUpdateDraft] = DraftStore()
    write = CalendarWriteUseCases(
        preview_event=PreviewEvent(create_drafts),
        create_event=CreateEvent(create_drafts, writer),
        preview_event_update=PreviewEventUpdate(writer, update_drafts),
        update_event=UpdateEvent(update_drafts, writer),
        preview_event_deletion=None,
        delete_event=None,
    )
    bundle = _bundle(InMemoryMailRepository(), InMemoryCalendarRepository(), FixedClock(FROZEN_NOW))
    return UseCases(
        list_unread_emails=bundle.list_unread_emails,
        search_emails=bundle.search_emails,
        get_email=bundle.get_email,
        list_todays_events=bundle.list_todays_events,
        list_upcoming_events=bundle.list_upcoming_events,
        list_emails=bundle.list_emails,
        count_emails=bundle.count_emails,
        top_senders=bundle.top_senders,
        list_largest_emails=bundle.list_largest_emails,
        list_folders=bundle.list_folders,
        list_custom_folders=bundle.list_custom_folders,
        resolve_folders=bundle.resolve_folders,
        search_contacts=bundle.search_contacts,
        calendar_write=write,
    )


def mail_draft_use_cases(
    writer: MailDraftWriter,
    reply_writer: ReplyDraftWriter | None = None,
    deleter: DraftDeleter | None = None,
) -> UseCases:
    """A bundle with the draft pairs wired to `writer` and `reply_writer`, each over its
    own fresh draft store, and calendar write left unset, as if only the draft flag were
    on. Omitting `reply_writer` gives the reply pair an empty mailbox; omitting `deleter`
    leaves the deletion pair unset, as if the draft delete flag were off.
    """
    drafts: DraftStore[MailDraftPreview] = DraftStore()
    reply_drafts: DraftStore[ReplyDraftPreview] = DraftStore()
    deletions: DraftStore[DraftDeletionPreview] = DraftStore()
    replies = reply_writer if reply_writer is not None else InMemoryReplyDraftWriter()
    return replace(
        mail_only_use_cases(InMemoryMailRepository()),
        mail_draft=MailDraftUseCases(
            preview_mail_draft=PreviewMailDraft(drafts),
            create_mail_draft=CreateMailDraft(drafts, writer),
            preview_reply_draft=PreviewReplyDraft(replies, reply_drafts),
            create_reply_draft=CreateReplyDraft(reply_drafts, replies),
            preview_draft_deletion=PreviewDraftDeletion(deleter, deletions) if deleter else None,
            delete_draft=DeleteDraft(deletions, deleter) if deleter else None,
        ),
    )


def folders_use_cases(repository: MailFolderRepository) -> UseCases:
    """A bundle for tests of the list_folders tool; every other side is empty."""
    bundle = _bundle(InMemoryMailRepository(), InMemoryCalendarRepository(), FixedClock(FROZEN_NOW))
    return UseCases(
        list_unread_emails=bundle.list_unread_emails,
        search_emails=bundle.search_emails,
        get_email=bundle.get_email,
        list_todays_events=bundle.list_todays_events,
        list_upcoming_events=bundle.list_upcoming_events,
        list_emails=bundle.list_emails,
        count_emails=bundle.count_emails,
        top_senders=bundle.top_senders,
        list_largest_emails=bundle.list_largest_emails,
        list_folders=ListFolders(repository),
        list_custom_folders=ListCustomFolders(repository),
        resolve_folders=ResolveFolders(repository),
        search_contacts=SearchContacts(InMemoryContactRepository()),
    )


def contacts_use_cases(repository: ContactRepository) -> UseCases:
    """A bundle for tests of the search_contacts tool; every other side is empty."""
    bundle = _bundle(InMemoryMailRepository(), InMemoryCalendarRepository(), FixedClock(FROZEN_NOW))
    return UseCases(
        list_unread_emails=bundle.list_unread_emails,
        search_emails=bundle.search_emails,
        get_email=bundle.get_email,
        list_todays_events=bundle.list_todays_events,
        list_upcoming_events=bundle.list_upcoming_events,
        list_emails=bundle.list_emails,
        count_emails=bundle.count_emails,
        top_senders=bundle.top_senders,
        list_largest_emails=bundle.list_largest_emails,
        list_folders=bundle.list_folders,
        list_custom_folders=bundle.list_custom_folders,
        resolve_folders=bundle.resolve_folders,
        search_contacts=SearchContacts(repository),
    )


def _bundle(
    mail: MailRepository,
    calendar: CalendarRepository,
    clock: Clock,
    folder_repository: MailFolderRepository | None = None,
) -> UseCases:
    if folder_repository is None:
        folder_repository = InMemoryMailFolderRepository()
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
        list_folders=ListFolders(folder_repository),
        list_custom_folders=ListCustomFolders(folder_repository),
        resolve_folders=ResolveFolders(folder_repository),
        search_contacts=SearchContacts(InMemoryContactRepository()),
    )
