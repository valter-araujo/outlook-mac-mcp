import pytest

from outlook_mac_mcp.application.resolve_folders import ResolveFolders
from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from outlook_mac_mcp.domain.errors import CustomFolderNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_folder_repository import InMemoryMailFolderRepository


class ExplodingFolderRepository(InMemoryMailFolderRepository):
    """Proves a resolution path never reaches the Graph tree walk: any call to
    list_custom fails the test rather than silently succeeding.
    """

    def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
        raise AssertionError("must not walk the mailbox tree")


def make_custom_folder(path: str, *, folder_id: str = "id") -> CustomMailFolder:
    return CustomMailFolder(
        folder_id=folder_id,
        display_name=path.rsplit("/", 1)[-1],
        path=path,
        unread_count=0,
        total_count=0,
    )


def test_resolves_a_well_known_name_to_itself_with_no_repository_call() -> None:
    resolved = ResolveFolders(ExplodingFolderRepository()).execute("archive")

    assert resolved.folder_ids == (FolderName.ARCHIVE,)
    assert resolved.echo == "archive"


def test_resolves_all_to_every_well_known_folder_with_no_repository_call() -> None:
    resolved = ResolveFolders(ExplodingFolderRepository()).execute("all")

    assert set(resolved.folder_ids) == set(FolderName)
    assert "all" not in resolved.echo.split(", ")


@pytest.mark.parametrize(
    ("raw", "expected"),
    [("INBOX", FolderName.INBOX), ("Archive", FolderName.ARCHIVE), ("ALL", None)],
)
def test_resolves_a_case_insensitive_well_known_match_with_no_repository_call(
    raw: str, expected: FolderName | None
) -> None:
    resolved = ResolveFolders(ExplodingFolderRepository()).execute(raw)

    if expected is None:
        assert set(resolved.folder_ids) == set(FolderName)
    else:
        assert resolved.folder_ids == (expected,)


@pytest.mark.parametrize(("raw", "expected_suggestion"), [("inbx", "inbox"), ("achive", "archive")])
def test_rejects_a_close_typo_of_a_well_known_name_with_a_suggestion_and_no_repository_call(
    raw: str, expected_suggestion: str
) -> None:
    with pytest.raises(InvalidRequestError, match=expected_suggestion):
        ResolveFolders(ExplodingFolderRepository()).execute(raw)


def test_leaves_a_genuine_custom_path_that_shares_letters_with_a_well_known_name_unaffected() -> (
    None
):
    repository = InMemoryMailFolderRepository()
    repository.set_custom(
        CustomFolderScan(
            folders=(make_custom_folder("Draft Ideas", folder_id="graph-id-1"),),
            depth_limit_reached=False,
            folder_limit_reached=False,
        )
    )

    resolved = ResolveFolders(repository).execute("Draft Ideas")

    assert resolved.folder_ids == ("graph-id-1",)
    assert resolved.echo == "Draft Ideas"


def test_all_still_means_only_the_five_well_known_folders_even_if_custom_ones_exist() -> None:
    repository = InMemoryMailFolderRepository()
    repository.set_custom(
        CustomFolderScan(
            folders=(make_custom_folder("Entrevistas/Work/AWS", folder_id="graph-id-1"),),
            depth_limit_reached=False,
            folder_limit_reached=False,
        )
    )

    resolved = ResolveFolders(repository).execute("all")

    assert set(resolved.folder_ids) == set(FolderName)
    assert "graph-id-1" not in resolved.folder_ids


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
