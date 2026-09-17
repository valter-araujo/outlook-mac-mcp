from dataclasses import dataclass

from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder


@dataclass(frozen=True, slots=True)
class CustomFolderScan:
    """What a walk over the mailbox's custom folder tree yields.

    Either flag being true means some part of the tree went unwalked: `folders` then
    holds fewer than the mailbox's real set. Both are reported as data rather than
    raised, the same way SenderScan reports a ceiling instead of raising.
    """

    folders: tuple[CustomMailFolder, ...]
    depth_limit_reached: bool
    folder_limit_reached: bool
