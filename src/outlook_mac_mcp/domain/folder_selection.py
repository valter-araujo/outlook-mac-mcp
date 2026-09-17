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


def ensure_well_formed_folders(folders: tuple[FolderName, ...]) -> None:
    """Every mail request that fans out over folders shares this one check, so a bad
    tuple is caught the same way regardless of which request built it.
    """
    if not folders:
        raise InvalidRequestError("folders must not be empty")
    if len(set(folders)) != len(folders):
        raise InvalidRequestError("folders must not repeat a folder")
