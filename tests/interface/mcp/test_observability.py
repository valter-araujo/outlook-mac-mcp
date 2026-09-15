import json
import logging

import pytest

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from outlook_mac_mcp.interface.mcp.observability import (
    DEFAULT_LOG_LEVEL,
    LOG_LEVEL_ENV_VAR,
    LOGGER_NAME,
    TOOL_CALL_EVENT,
    configure_logging,
    observed_tool_call,
)

A_TOOL = "list_unread_emails"
SECRET_MESSAGE = "mailbox of ceo@example.com held subject 'Board pay'"


def read_records(capsys: pytest.CaptureFixture[str]) -> list[dict[str, object]]:
    """Assert nothing reached stdout, then parse each stderr line as one JSON record."""
    captured = capsys.readouterr()
    assert captured.out == ""
    return [json.loads(line) for line in captured.err.splitlines() if line]


def test_writes_one_json_record_per_call_to_stderr(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()

    with observed_tool_call(A_TOOL):
        pass

    assert len(read_records(capsys)) == 1


def test_records_the_tool_name_and_a_correlation_id(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()

    with observed_tool_call(A_TOOL):
        pass

    record = read_records(capsys)[0]
    assert record["event"] == TOOL_CALL_EVENT
    assert record["tool_name"] == A_TOOL
    assert record["correlation_id"]


def test_gives_each_call_a_different_correlation_id(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()

    with observed_tool_call(A_TOOL):
        pass
    with observed_tool_call(A_TOOL):
        pass

    first, second = read_records(capsys)
    assert first["correlation_id"] != second["correlation_id"]


def test_records_success_the_item_count_and_a_duration(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()

    with observed_tool_call(A_TOOL) as outcome:
        outcome.item_count = 7

    record = read_records(capsys)[0]
    assert record["succeeded"] is True
    assert record["error_type"] == ""
    assert record["item_count"] == 7
    assert isinstance(record["duration_ms"], float)


def test_records_the_type_of_a_project_error_and_re_raises_it(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()

    with pytest.raises(InvalidRequestError), observed_tool_call(A_TOOL):
        raise InvalidRequestError(SECRET_MESSAGE)

    record = read_records(capsys)[0]
    assert record["succeeded"] is False
    assert record["error_type"] == "InvalidRequestError"


def test_never_writes_the_error_message_into_the_log(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()

    with pytest.raises(GraphRequestError), observed_tool_call(A_TOOL):
        raise GraphRequestError(SECRET_MESSAGE)

    logged = json.dumps(read_records(capsys)[0])
    assert "ceo@example.com" not in logged
    assert "Board pay" not in logged


def test_records_an_unexpected_error_without_catching_it(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()

    with pytest.raises(ZeroDivisionError), observed_tool_call(A_TOOL):
        raise ZeroDivisionError(SECRET_MESSAGE)

    record = read_records(capsys)[0]
    assert record["succeeded"] is False
    assert record["error_type"] == "UnknownError"


def test_defaults_to_info_level(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(LOG_LEVEL_ENV_VAR, raising=False)

    configure_logging()

    assert logging.getLogger(LOGGER_NAME).level == logging.getLevelName(DEFAULT_LOG_LEVEL)


def test_takes_its_level_from_the_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(LOG_LEVEL_ENV_VAR, "WARNING")

    configure_logging()

    assert logging.getLogger(LOGGER_NAME).level == logging.WARNING
