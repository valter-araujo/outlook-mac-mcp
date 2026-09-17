from outlook_mac_mcp.application.list_custom_folders import ListCustomFolders
from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from tests.fakes.in_memory_mail_folder_repository import InMemoryMailFolderRepository


def make_folder(
    path: str, folder_id: str = "id", *, unread: int = 0, total: int = 0
) -> CustomMailFolder:
    return CustomMailFolder(
        folder_id=folder_id,
        display_name=path.rsplit("/", 1)[-1],
        path=path,
        unread_count=unread,
        total_count=total,
    )


def test_returns_whatever_the_repository_scanned() -> None:
    repository = InMemoryMailFolderRepository()
    scan = CustomFolderScan(
        folders=(make_folder("Candidaturas"),),
        depth_limit_reached=False,
        folder_limit_reached=False,
    )
    repository.set_custom(scan)

    result = ListCustomFolders(repository).execute()

    assert result == scan


def test_returns_empty_when_the_repository_holds_no_custom_folders() -> None:
    result = ListCustomFolders(InMemoryMailFolderRepository()).execute()

    assert result.folders == ()
    assert result.depth_limit_reached is False
    assert result.folder_limit_reached is False


def test_passes_its_caps_through_to_the_repository() -> None:
    passed_args: list[tuple[int, int]] = []

    class RecordingRepository(InMemoryMailFolderRepository):
        def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
            passed_args.append((max_depth, max_folders))
            return super().list_custom(max_depth, max_folders)

    ListCustomFolders(RecordingRepository(), max_depth=3, max_folders=7).execute()

    assert passed_args == [(3, 7)]
