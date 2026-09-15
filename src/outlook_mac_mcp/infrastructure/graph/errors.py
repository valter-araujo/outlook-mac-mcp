from outlook_mac_mcp.domain.errors import OutlookMcpError


class GraphError(OutlookMcpError):
    """Base for every failure raised by the Microsoft Graph adapter.

    Graph-specific errors live here and not in `domain/` so the domain stays free
    of any knowledge of Microsoft, while the MCP boundary still catches a single
    `OutlookMcpError`.
    """


class ConfigurationError(GraphError):
    """The adapter is missing configuration it cannot invent, such as the client id."""


class TokenCacheError(GraphError):
    """The Keychain could not be read or written."""


class NotAuthenticatedError(GraphError):
    """No usable credential is cached; the user must complete a sign-in first."""


class AuthenticationError(GraphError):
    """The identity platform refused to issue a token."""


class GraphRequestError(GraphError):
    """Graph answered with a non-success status."""


class GraphResponseError(GraphError):
    """Graph answered successfully but the payload was not shaped as expected."""


class UnsupportedHostError(GraphError):
    """A request was about to leave the one host this adapter is allowed to reach."""
