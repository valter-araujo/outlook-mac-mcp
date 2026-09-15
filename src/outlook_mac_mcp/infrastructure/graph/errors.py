from outlook_mac_mcp.domain.errors import OutlookMcpError


class GraphError(OutlookMcpError):
    """Base for every failure raised by the Microsoft Graph adapter.

    Graph-specific errors live here and not in `domain/` so the domain stays free
    of any knowledge of Microsoft, while the MCP boundary still catches a single
    `OutlookMcpError`.
    """


class TokenCacheError(GraphError):
    """The Keychain could not be read or written."""


class NotAuthenticatedError(GraphError):
    """No usable credential is cached; the user must complete a sign-in first."""


class AuthenticationError(GraphError):
    """The identity platform refused to issue a token."""


class GraphRequestError(GraphError):
    """Graph answered with a non-success status.

    Carries the status and Graph's own error code so callers can tell an expected
    condition, such as a message that does not exist or an id that is not well formed,
    from an infrastructure failure, without parsing the message text.
    """

    def __init__(self, message: str, status_code: int, error_code: str) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.error_code = error_code


class GraphResponseError(GraphError):
    """Graph answered successfully but the payload was not shaped as expected."""


class UnsupportedHostError(GraphError):
    """A request was about to leave the one host this adapter is allowed to reach."""
