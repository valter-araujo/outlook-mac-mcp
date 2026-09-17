import re
from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.folder_selection import ensure_well_formed_folders

# One @, no whitespace, a dot in the domain, and no single quote: the quote is the OData
# string delimiter, and refusing it is simpler and safer than escaping it.
SENDER_ADDRESS_PATTERN = re.compile(r"^[^@\s']+@[^@\s']+\.[^@\s']+$")
MAX_SENDER_ADDRESS_LENGTH = 320


@dataclass(frozen=True, slots=True)
class EmailFilters:
    """Every optional restriction a mail listing or count can apply, combinable at will.

    `received_after` is inclusive and `received_before` exclusive, so consecutive ranges
    neither overlap nor leave a gap. The sender is matched exactly on the address.

    `folders` holds more than one well-known folder only for the `all` selection: the
    adapter runs the same query against each and merges the results, never splitting a
    single caller-named folder into more than one.
    """

    folders: tuple[str, ...] = (FolderName.INBOX,)
    is_read: bool | None = None
    sender: str | None = None
    received_after: datetime | None = None
    received_before: datetime | None = None
    has_attachments: bool | None = None

    def __post_init__(self) -> None:
        ensure_well_formed_folders(self.folders)
        if self.sender is not None:
            _ensure_sender_is_an_address(self.sender)
        for bound in (self.received_after, self.received_before):
            if bound is not None and bound.tzinfo is None:
                raise InvalidRequestError("received_after and received_before need a time zone")
        if (
            self.received_after is not None
            and self.received_before is not None
            and not self.received_after < self.received_before
        ):
            raise InvalidRequestError("received_after must be before received_before")

    def is_empty(self) -> bool:
        return all(
            value is None
            for value in (
                self.is_read,
                self.sender,
                self.received_after,
                self.received_before,
                self.has_attachments,
            )
        )


def _ensure_sender_is_an_address(sender: str) -> None:
    if len(sender) > MAX_SENDER_ADDRESS_LENGTH or not SENDER_ADDRESS_PATTERN.match(sender):
        raise InvalidRequestError("sender must be a single email address")
