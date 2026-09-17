from enum import StrEnum

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName


class FolderSelection(StrEnum):
    """Every folder-scope value a mail tool accepts: one well-known folder, or `all`,
    which fans out to every well-known folder and merges the results.

    Values mirror FolderName's so the two never drift apart; `all` is the one addition.
    """

    INBOX = FolderName.INBOX.value
    ARCHIVE = FolderName.ARCHIVE.value
    JUNK = FolderName.JUNK.value
    SENT = FolderName.SENT.value
    DRAFTS = FolderName.DRAFTS.value
    ALL = "all"

    def to_folders(self) -> tuple[FolderName, ...]:
        if self is FolderSelection.ALL:
            return tuple(FolderName)
        return (FolderName(self.value),)

    def describe_folders(self) -> str:
        """Every folder this selection actually covers, as a readable list -- not just
        the selector itself. Two `all` results stay comparable at a glance, including if
        the set of well-known folders ever changes, without having to read the code to
        know what `all` meant at the time.
        """
        return ", ".join(folder.value for folder in self.to_folders())


def ensure_well_formed_folders(folders: tuple[str, ...]) -> None:
    """Every mail request that fans out over folders shares this one check, so a bad
    tuple is caught the same way regardless of which request built it.

    `folders` holds Graph folder identifiers generally, not only `FolderName` members:
    a resolved custom folder's Graph id passes through here too.
    """
    if not folders:
        raise InvalidRequestError("folders must not be empty")
    if len(set(folders)) != len(folders):
        raise InvalidRequestError("folders must not repeat a folder")
