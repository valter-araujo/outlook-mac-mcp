class OutlookMcpError(Exception):
    """Base for every error raised by this project.

    Callers at the boundary (the MCP layer) catch only this type, so any
    failure that reaches the user is one we chose to surface.
    """


class InvalidRequestError(OutlookMcpError):
    """A use case received a request that violates a domain rule."""


class EmailNotFoundError(OutlookMcpError):
    """The mailbox holds no message with the requested id."""


class EventNotFoundError(OutlookMcpError):
    """The calendar holds no event with the requested id."""


class DraftNotFoundError(OutlookMcpError):
    """No draft is waiting under the given token: it was never issued, already used,
    or issued by a server process that has since restarted.
    """


class EventWriteUnconfirmedError(OutlookMcpError):
    """A create or update reached the calendar and succeeded, but the response
    describing the result could not be read afterward.

    The write already happened; only confirming its result failed. Raising this
    instead of the underlying read error matters because that error, on its own,
    reads exactly like the write itself failed — the one thing it must never be
    mistaken for, since a caller who believes nothing happened has every reason to
    retry, and a retry after a write that actually succeeded creates a duplicate.
    """
