import json
from collections.abc import Callable

import pytest

from outlook_mac_mcp import cli
from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.bootstrap import UseCases
from outlook_mac_mcp.infrastructure.graph.authentication import (
    DeviceCodeAuthenticator,
    DeviceCodePrompt,
)
from outlook_mac_mcp.infrastructure.graph.errors import ConfigurationError
from outlook_mac_mcp.infrastructure.settings import CLIENT_ID_ENV_VAR, TIMEZONE_ENV_VAR
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

A_PROMPT = DeviceCodePrompt(
    user_code="ABCD-EFGH",
    verification_uri="https://microsoft.com/devicelogin",
    expires_in_seconds=900,
)


class RecordingAuthenticator:
    """Stands in for the real authenticator so no browser or Keychain is involved."""

    def __init__(self, error: Exception | None = None) -> None:
        self.error = error
        self.signed_in = False

    def sign_in(self, show_prompt: Callable[[DeviceCodePrompt], None]) -> None:
        if self.error is not None:
            raise self.error
        show_prompt(A_PROMPT)
        self.signed_in = True


@pytest.fixture(autouse=True)
def configured_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(CLIENT_ID_ENV_VAR, "a-client-id")
    monkeypatch.setenv(TIMEZONE_ENV_VAR, "Europe/Lisbon")


def install_authenticator(
    monkeypatch: pytest.MonkeyPatch, authenticator: RecordingAuthenticator
) -> None:
    monkeypatch.setattr(
        DeviceCodeAuthenticator,
        "from_settings",
        classmethod(lambda cls, settings: authenticator),
    )


def test_defaults_to_serving_when_no_command_is_given() -> None:
    assert cli._parse_command([]) == cli.SERVE_COMMAND


def test_selects_the_sign_in_command() -> None:
    assert cli._parse_command([cli.SIGN_IN_COMMAND]) == cli.SIGN_IN_COMMAND


def test_rejects_an_unknown_command() -> None:
    with pytest.raises(SystemExit):
        cli._parse_command(["launch-missiles"])


def test_sign_in_runs_the_device_code_flow(monkeypatch: pytest.MonkeyPatch) -> None:
    authenticator = RecordingAuthenticator()
    install_authenticator(monkeypatch, authenticator)

    assert cli.main([cli.SIGN_IN_COMMAND]) == cli.EXIT_SUCCESS
    assert authenticator.signed_in is True


def test_sign_in_prints_the_code_to_stderr_and_never_to_stdout(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    install_authenticator(monkeypatch, RecordingAuthenticator())

    cli.main([cli.SIGN_IN_COMMAND])

    captured = capsys.readouterr()
    assert captured.out == ""
    assert "ABCD-EFGH" in captured.err
    assert "https://microsoft.com/devicelogin" in captured.err


def test_sign_in_reports_a_project_error_on_stderr_and_fails(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    install_authenticator(
        monkeypatch, RecordingAuthenticator(ConfigurationError("OUTLOOK_MCP_CLIENT_ID is not set"))
    )

    exit_code = cli.main([cli.SIGN_IN_COMMAND])

    captured = capsys.readouterr()
    assert exit_code == cli.EXIT_FAILURE
    assert captured.out == ""
    assert "OUTLOOK_MCP_CLIENT_ID is not set" in captured.err


def test_sign_in_fails_before_the_flow_when_the_time_zone_is_invalid(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    authenticator = RecordingAuthenticator()
    install_authenticator(monkeypatch, authenticator)
    monkeypatch.setenv(TIMEZONE_ENV_VAR, "Mars/Olympus_Mons")

    exit_code = cli.main([cli.SIGN_IN_COMMAND])

    assert exit_code == cli.EXIT_FAILURE
    assert authenticator.signed_in is False
    assert "OUTLOOK_MCP_TIMEZONE" in capsys.readouterr().err


def test_serving_logs_the_resolved_time_zone_to_stderr(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    install_authenticator(monkeypatch, RecordingAuthenticator())
    monkeypatch.setattr(cli, "build_use_cases", lambda settings: _use_cases())
    monkeypatch.setattr(
        cli, "build_server", lambda list_unread_emails, search_emails, get_email: _ServerSpy([])
    )

    cli.main([cli.SERVE_COMMAND])

    captured = capsys.readouterr()
    assert captured.out == ""
    startup = next(json.loads(line) for line in captured.err.splitlines() if "startup" in line)
    assert startup["event"] == "startup"
    assert startup["timezone"] == "Europe/Lisbon"


def _use_cases() -> UseCases:
    repository = InMemoryMailRepository()
    return UseCases(
        list_unread_emails=ListUnreadEmails(repository),
        search_emails=SearchEmails(repository),
        get_email=GetEmail(repository),
    )


def test_serving_does_not_start_a_device_code_flow(monkeypatch: pytest.MonkeyPatch) -> None:
    authenticator = RecordingAuthenticator()
    install_authenticator(monkeypatch, authenticator)
    started: list[bool] = []
    monkeypatch.setattr(cli, "build_use_cases", lambda settings: _use_cases())
    monkeypatch.setattr(
        cli,
        "build_server",
        lambda list_unread_emails, search_emails, get_email: _ServerSpy(started),
    )

    assert cli.main([cli.SERVE_COMMAND]) == cli.EXIT_SUCCESS
    assert started == [True]
    assert authenticator.signed_in is False


class _ServerSpy:
    def __init__(self, started: list[bool]) -> None:
        self._started = started

    def run(self) -> None:
        self._started.append(True)
