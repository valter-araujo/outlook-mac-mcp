import logging

LOGGER_NAME = "outlook_mac_mcp"


def project_logger() -> logging.Logger:
    """The one logger every layer writes to, so a single handler decides where it goes."""
    return logging.getLogger(LOGGER_NAME)
