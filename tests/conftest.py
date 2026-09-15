import logging
from collections.abc import Iterator

import pytest

from outlook_mac_mcp.interface.mcp.observability import LOGGER_NAME


@pytest.fixture
def anyio_backend() -> str:
    """Run async tests on asyncio only; the MCP SDK is exercised the way it ships."""
    return "asyncio"


@pytest.fixture(autouse=True)
def restore_project_logger() -> Iterator[None]:
    """Keep logging configuration from leaking between tests.

    `configure_logging` binds a handler to whatever `sys.stderr` is at the time. Without
    this, a test that captures stderr would leave a closed stream behind for the next one.
    """
    logger = logging.getLogger(LOGGER_NAME)
    handlers = list(logger.handlers)
    level = logger.level
    propagate = logger.propagate
    yield
    logger.handlers = handlers
    logger.setLevel(level)
    logger.propagate = propagate
