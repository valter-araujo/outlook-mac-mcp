class OutlookMcpError(Exception):
    """Base for every error raised by this project.

    Callers at the boundary (the MCP layer) catch only this type, so any
    failure that reaches the user is one we chose to surface.
    """


class InvalidRequestError(OutlookMcpError):
    """A use case received a request that violates a domain rule."""


class EmailNotFoundError(OutlookMcpError):
    """The mailbox holds no message with the requested id."""


class DraftNotFoundError(OutlookMcpError):
    """No draft is waiting under the given token: it was never issued, already used,
    or issued by a server process that has since restarted.
    """
