import pytest

from outlook_mac_mcp.application.resolve_folders import ResolveFolders
from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from outlook_mac_mcp.domain.errors import CustomFolderNotFoundError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_folder_repository import InMemoryMailFolderRepository


def make_custom_folder(path: str, *, folder_id: str = "id") -> CustomMailFolder:
    return CustomMailFolder(
        folder_id=folder_id,
        display_name=path.rsplit("/", 1)[-1],
        path=path,
        unread_count=0,
        total_count=0,
    )


def test_resolves_a_well_known_name_to_itself_with_no_repository_call() -> None:
    class ExplodingRepository(InMemoryMailFolderRepository):
        def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
            raise AssertionError("must not be called for a well-known folder")

    resolved = ResolveFolders(ExplodingRepository()).execute("archive")

    assert resolved.folder_ids == (FolderName.ARCHIVE,)
    assert resolved.echo == "archive"


def test_resolves_all_to_every_well_known_folder_with_no_repository_call() -> None:
    class ExplodingRepository(InMemoryMailFolderRepository):
        def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
            raise AssertionError("must not be called for all")

    resolved = ResolveFolders(ExplodingRepository()).execute("all")

    assert set(resolved.folder_ids) == set(FolderName)
    assert "all" not in resolved.echo.split(", ")


def test_resolves_a_custom_path_to_its_graph_id() -> None:
    repository = InMemoryMailFolderRepository()
    repository.set_custom(
        CustomFolderScan(
            folders=(make_custom_folder("Entrevistas/Work/AWS", folder_id="graph-id-1"),),
            depth_limit_reached=False,
            folder_limit_reached=False,
        )
    )

    resolved = ResolveFolders(repository).execute("Entrevistas/Work/AWS")

    assert resolved.folder_ids == ("graph-id-1",)
    assert resolved.echo == "Entrevistas/Work/AWS"


def test_disambiguates_two_folders_sharing_a_display_name_by_full_path() -> None:
    repository = InMemoryMailFolderRepository()
    repository.set_custom(
        CustomFolderScan(
            folders=(
                make_custom_folder("Candidaturas/2026", folder_id="f1"),
                make_custom_folder("Projetos/2026", folder_id="f2"),
            ),
            depth_limit_reached=False,
            folder_limit_reached=False,
        )
    )

    resolved = ResolveFolders(repository).execute("Projetos/2026")

    assert resolved.folder_ids == ("f2",)


def test_raises_a_clear_error_naming_what_was_searched_when_no_path_matches() -> None:
    repository = InMemoryMailFolderRepository()
    resolve = ResolveFolders(repository, max_depth=3, max_folders=50)

    with pytest.raises(CustomFolderNotFoundError, match=r"Nowhere.*3.*50"):
        resolve.execute("Nowhere")


def test_passes_its_caps_through_to_the_repository_walk() -> None:
    passed_args: list[tuple[int, int]] = []

    class RecordingRepository(InMemoryMailFolderRepository):
        def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
            passed_args.append((max_depth, max_folders))
            return super().list_custom(max_depth, max_folders)

    with pytest.raises(CustomFolderNotFoundError):
        ResolveFolders(RecordingRepository(), max_depth=4, max_folders=9).execute("Nowhere")

    assert passed_args == [(4, 9)]
