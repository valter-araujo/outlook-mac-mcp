from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class CustomMailFolder:
    """A user-created folder discovered by walking the mailbox's folder tree.

    `path` is built from parent traversal (e.g. "Candidaturas/2026") and is what
    disambiguates two folders that share a display name under different parents;
    `folder_id` is Graph's own id for the folder.
    """

    folder_id: str
    display_name: str
    path: str
    unread_count: int
    total_count: int
