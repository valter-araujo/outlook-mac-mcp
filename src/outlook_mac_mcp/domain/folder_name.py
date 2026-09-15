from enum import StrEnum


class FolderName(StrEnum):
    """Well-known mailbox folders.

    Values are the Graph API well-known names so the adapter needs no mapping.
    Arbitrary user folders are out of scope for v1.
    """

    INBOX = "inbox"
    ARCHIVE = "archive"
    JUNK = "junkemail"
    SENT = "sentitems"
    DRAFTS = "drafts"
