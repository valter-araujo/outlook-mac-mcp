import json
import logging
import os
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import UTC, datetime
from time import perf_counter
from uuid import uuid4

from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.logger import project_logger

LOG_LEVEL_ENV_VAR = "OUTLOOK_MCP_LOG_LEVEL"
DEFAULT_LOG_LEVEL = "INFO"
TOOL_CALL_EVENT = "tool_call"
STARTUP_EVENT = "startup"
UNKNOWN_ERROR_TYPE = "UnknownError"
NO_ERROR_TYPE = ""


@dataclass(frozen=True, slots=True)
class ToolCallRecord:
    tool_name: str
    correlation_id: str
    duration_ms: float
    succeeded: bool
    error_type: str
    item_count: int


class ToolCallOutcome:
    """Mutable on purpose: the item count is only known after the use case returns."""

    def __init__(self) -> None:
        self.item_count = 0


class JsonFormatter(logging.Formatter):
    """Renders one log record as a single JSON object.

    Only the fields the caller put in `extra["fields"]` are emitted, so mailbox content
    cannot reach the log by accident: there is no formatting of arbitrary arguments.
    """

    def format(self, record: logging.LogRecord) -> str:
        payload = {
            "timestamp": datetime.fromtimestamp(record.created, tz=UTC).isoformat(),
            "level": record.levelname,
            "event": record.getMessage(),
        }
        fields = getattr(record, "fields", None)
        if isinstance(fields, dict):
            payload.update(fields)
        return json.dumps(payload)


def configure_logging() -> None:
    """Send structured logs to stderr only; stdout belongs to the MCP transport."""
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(JsonFormatter())
    logger = project_logger()
    logger.setLevel(os.environ.get(LOG_LEVEL_ENV_VAR, DEFAULT_LOG_LEVEL))
    logger.handlers = [handler]
    logger.propagate = False


def record_startup(timezone_name: str) -> None:
    """Name the zone every calendar time will be expressed in; a zone is metadata, not content."""
    project_logger().info(STARTUP_EVENT, extra={"fields": {"timezone": timezone_name}})


@contextmanager
def observed_tool_call(tool_name: str) -> Iterator[ToolCallOutcome]:
    """Record one tool call: who, how long, whether it worked and how much it returned.

    Project errors are caught only to name their type before re-raising; the message is
    deliberately left out, because a Graph error can quote the query or mailbox content.
    Anything else is never caught, so `finally` is what guarantees every call is recorded.
    """
    outcome = ToolCallOutcome()
    correlation_id = uuid4().hex
    started_at = perf_counter()
    succeeded = False
    error_type = UNKNOWN_ERROR_TYPE
    try:
        yield outcome
        succeeded = True
        error_type = NO_ERROR_TYPE
    except OutlookMcpError as error:
        error_type = type(error).__name__
        raise
    finally:
        _write(
            ToolCallRecord(
                tool_name=tool_name,
                correlation_id=correlation_id,
                duration_ms=round((perf_counter() - started_at) * 1000, 3),
                succeeded=succeeded,
                error_type=error_type,
                item_count=outcome.item_count,
            )
        )


def _write(record: ToolCallRecord) -> None:
    level = logging.INFO if record.succeeded else logging.ERROR
    project_logger().log(level, TOOL_CALL_EVENT, extra={"fields": asdict(record)})
