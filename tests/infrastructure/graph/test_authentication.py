import os
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.infrastructure.graph.authentication import (
    CALENDAR_WRITE_SCOPE,
    READ_SCOPES,
    DeviceCodeAuthenticator,
    DeviceCodePrompt,
    required_scopes,
)
from outlook_mac_mcp.infrastructure.graph.errors import (
    AuthenticationError,
    NotAuthenticatedError,
)
from outlook_mac_mcp.infrastructure.settings import Settings, load_settings

AN_ACCOUNT = {"username": "someone@example.com"}
RESERVED_SCOPES = ("offline_access", "openid", "profile")
A_DEVICE_FLOW = {
    "user_code": "ABCD-EFGH",
    "verification_uri": "https://microsoft.com/devicelogin",
    "expires_in": 900,
    "device_code": "opaque",
}


@dataclass
class FakeMsalApplication:
    """Stands in for msal.PublicClientApplication, the identity-platform boundary."""

    accounts: list[object] = field(default_factory=list)
    silent_response: Mapping[str, Any] | None = None
    device_flow: Mapping[str, Any] = field(default_factory=dict)
    device_flow_response: Mapping[str, Any] = field(default_factory=dict)
    requested_scopes: list[str] = field(default_factory=list)

    def get_accounts(self) -> Sequence[object]:
        return self.accounts

    def acquire_token_silent(self, scopes: list[str], account: object) -> Mapping[str, Any] | None:
        self.requested_scopes = scopes
        return self.silent_response

    def initiate_device_flow(self, scopes: list[str]) -> Mapping[str, Any]:
        self.requested_scopes = scopes
        return self.device_flow

    def acquire_token_by_device_flow(self, flow: Mapping[str, Any]) -> Mapping[str, Any]:
        return self.device_flow_response


@dataclass
class FakeTokenCache:
    saves: int = 0

    def save_if_changed(self) -> None:
        self.saves += 1


def test_returns_the_access_token_held_in_the_cache() -> None:
    application = FakeMsalApplication(
        accounts=[AN_ACCOUNT], silent_response={"access_token": "a-token"}
    )
    authenticator = DeviceCodeAuthenticator(application, FakeTokenCache())

    assert authenticator.get_access_token() == "a-token"


def test_raises_sign_in_required_when_no_account_is_cached() -> None:
    authenticator = DeviceCodeAuthenticator(FakeMsalApplication(), FakeTokenCache())

    with pytest.raises(NotAuthenticatedError):
        authenticator.get_access_token()


def test_raises_sign_in_required_when_the_token_cannot_be_refreshed_silently() -> None:
    application = FakeMsalApplication(accounts=[AN_ACCOUNT], silent_response=None)
    authenticator = DeviceCodeAuthenticator(application, FakeTokenCache())

    with pytest.raises(NotAuthenticatedError):
        authenticator.get_access_token()


def test_raises_when_the_identity_platform_reports_an_error() -> None:
    application = FakeMsalApplication(
        accounts=[AN_ACCOUNT],
        silent_response={"error": "invalid_grant", "error_description": "expired"},
    )
    authenticator = DeviceCodeAuthenticator(application, FakeTokenCache())

    with pytest.raises(AuthenticationError):
        authenticator.get_access_token()


def test_persists_the_cache_after_a_silent_refresh() -> None:
    application = FakeMsalApplication(
        accounts=[AN_ACCOUNT], silent_response={"access_token": "a-token"}
    )
    token_cache = FakeTokenCache()

    DeviceCodeAuthenticator(application, token_cache).get_access_token()

    assert token_cache.saves == 1


def settings_with_calendar_write(enabled: bool) -> Settings:
    return Settings(
        client_id="a-client-id", timezone=ZoneInfo("UTC"), calendar_write_enabled=enabled
    )


@pytest.mark.parametrize("reserved_scope", RESERVED_SCOPES)
def test_never_requests_scopes_that_msal_reserves(reserved_scope: str) -> None:
    assert reserved_scope not in required_scopes(settings_with_calendar_write(True))


def test_requests_only_read_scopes_when_calendar_write_is_off() -> None:
    scopes = required_scopes(settings_with_calendar_write(False))

    assert scopes == READ_SCOPES
    assert CALENDAR_WRITE_SCOPE not in scopes


def test_adds_the_calendar_write_scope_when_the_flag_is_on() -> None:
    scopes = required_scopes(settings_with_calendar_write(True))

    assert set(READ_SCOPES) < set(scopes)
    assert CALENDAR_WRITE_SCOPE in scopes


def test_asks_the_identity_platform_for_exactly_the_scopes_it_was_given() -> None:
    application = FakeMsalApplication(accounts=[AN_ACCOUNT], silent_response={"access_token": "t"})
    authenticator = DeviceCodeAuthenticator(
        application, FakeTokenCache(), (*READ_SCOPES, CALENDAR_WRITE_SCOPE)
    )

    authenticator.get_access_token()

    assert application.requested_scopes == [*READ_SCOPES, CALENDAR_WRITE_SCOPE]


def test_asks_for_read_scopes_only_by_default() -> None:
    application = FakeMsalApplication(accounts=[AN_ACCOUNT], silent_response={"access_token": "t"})

    DeviceCodeAuthenticator(application, FakeTokenCache()).get_access_token()

    assert application.requested_scopes == list(READ_SCOPES)


def test_sign_in_shows_the_device_code_prompt() -> None:
    application = FakeMsalApplication(
        device_flow=A_DEVICE_FLOW, device_flow_response={"access_token": "a-token"}
    )
    shown: list[DeviceCodePrompt] = []

    DeviceCodeAuthenticator(application, FakeTokenCache()).sign_in(shown.append)

    assert shown == [
        DeviceCodePrompt(
            user_code="ABCD-EFGH",
            verification_uri="https://microsoft.com/devicelogin",
            expires_in_seconds=900,
        )
    ]


def test_sign_in_persists_the_cache_once_the_flow_completes() -> None:
    application = FakeMsalApplication(
        device_flow=A_DEVICE_FLOW, device_flow_response={"access_token": "a-token"}
    )
    token_cache = FakeTokenCache()

    DeviceCodeAuthenticator(application, token_cache).sign_in(lambda prompt: None)

    assert token_cache.saves == 1


def test_sign_in_raises_when_the_device_flow_cannot_start() -> None:
    application = FakeMsalApplication(device_flow={"error": "invalid_client"})
    authenticator = DeviceCodeAuthenticator(application, FakeTokenCache())

    with pytest.raises(AuthenticationError):
        authenticator.sign_in(lambda prompt: None)


def test_sign_in_raises_when_the_user_never_completes_the_flow() -> None:
    application = FakeMsalApplication(
        device_flow=A_DEVICE_FLOW,
        device_flow_response={"error": "expired_token", "error_description": "code expired"},
    )
    authenticator = DeviceCodeAuthenticator(application, FakeTokenCache())

    with pytest.raises(AuthenticationError):
        authenticator.sign_in(lambda prompt: None)


@pytest.mark.integration
def test_signs_in_against_the_real_identity_platform() -> None:
    authenticator = DeviceCodeAuthenticator.from_settings(load_settings(os.environ))

    authenticator.sign_in(lambda prompt: print(prompt, file=sys.stderr))

    assert authenticator.get_access_token()
