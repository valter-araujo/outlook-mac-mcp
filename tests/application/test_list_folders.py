from outlook_mac_mcp.application.list_folders import ListFolders
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder
from tests.fakes.in_memory_mail_folder_repository import InMemoryMailFolderRepository


def make_folder(name: FolderName, *, unread: int = 0, total: int = 0) -> MailFolder:
    return MailFolder(
        well_known_name=name,
        display_name=name.value.title(),
        unread_count=unread,
        total_count=total,
    )


def test_returns_every_folder_the_repository_holds() -> None:
    repository = InMemoryMailFolderRepository()
    repository.add(make_folder(FolderName.INBOX, unread=3, total=71))
    repository.add(make_folder(FolderName.ARCHIVE))

    folders = ListFolders(repository).execute()

    assert [folder.well_known_name for folder in folders] == [FolderName.INBOX, FolderName.ARCHIVE]
    assert folders[0].unread_count == 3
    assert folders[0].total_count == 71


def test_returns_folders_in_declaration_order_regardless_of_add_order() -> None:
    repository = InMemoryMailFolderRepository()
    repository.add(make_folder(FolderName.DRAFTS))
    repository.add(make_folder(FolderName.INBOX))
    repository.add(make_folder(FolderName.ARCHIVE))

    folders = ListFolders(repository).execute()

    assert [folder.well_known_name for folder in folders] == [
        FolderName.INBOX,
        FolderName.ARCHIVE,
        FolderName.DRAFTS,
    ]


def test_returns_empty_when_the_repository_holds_nothing() -> None:
    assert ListFolders(InMemoryMailFolderRepository()).execute() == ()
