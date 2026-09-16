from dataclasses import dataclass

from outlook_mac_mcp.domain.folder_name import FolderName


@dataclass(frozen=True, slots=True)
class MailFolder:
    """One of the mailbox's well-known folders, with its live item counts.

    `well_known_name` is the folder's well-known name, which Graph documents as a
    substitute for its opaque id and which this project already uses to address every
    folder; `display_name` is the mailbox's own, possibly localized, label for it.
    """

    well_known_name: FolderName
    display_name: str
    unread_count: int
    total_count: int
