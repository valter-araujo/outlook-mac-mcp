from outlook_mac_mcp.domain.errors import OutlookMcpError


class ConfigurationError(OutlookMcpError):
    """The process is missing configuration it cannot invent, or was given a bad value.

    Sits beside the Graph hierarchy rather than inside it: a missing client id or an
    unknown time zone is a fault of the environment, not of a Microsoft call, and the
    settings loader that raises it never touches Graph.
    """
