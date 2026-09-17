from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.folder_selection import FolderSelection


def test_a_single_folder_expands_to_just_itself() -> None:
    assert FolderSelection.ARCHIVE.to_folders() == (FolderName.ARCHIVE,)


def test_all_expands_to_every_well_known_folder() -> None:
    assert set(FolderSelection.ALL.to_folders()) == set(FolderName)


def test_a_single_folder_describes_as_its_own_name() -> None:
    assert FolderSelection.ARCHIVE.describe_folders() == "archive"


def test_all_describes_as_the_real_folder_names_not_the_word_all() -> None:
    """Two `all` results must stay comparable without reading the code: the output
    names exactly which folders were covered, not the literal selector.
    """
    described = FolderSelection.ALL.describe_folders()

    assert "all" not in described.split(", ")
    assert set(described.split(", ")) == {folder.value for folder in FolderName}
