from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from typing import Any, Protocol

from msal import PublicClientApplication

from outlook_mac_mcp.infrastructure.graph.errors import (
    AuthenticationError,
    NotAuthenticatedError,
)
from outlook_mac_mcp.infrastructure.graph.token_cache import KeyringTokenCache, TokenCache
from outlook_mac_mcp.infrastructure.settings import Settings

CONSUMERS_AUTHORITY = "https://login.microsoftonline.com/consumers"

# The reserved scopes (offline_access, openid, profile) are absent on purpose: MSAL
# appends them itself and rejects them when requested explicitly. The Entra
# registration still consents to offline_access, which is what allows the refresh
# token that keeps this cache useful between runs.
SCOPES: tuple[str, ...] = ("Mail.Read", "Calendars.Read", "Contacts.Read", "User.Read")

SIGN_IN_REQUIRED_MESSAGE = "no usable cached credential; complete the device-code sign-in first"


@dataclass(frozen=True, slots=True)
class DeviceCodePrompt:
    """What the user must be shown to complete a device-code sign-in."""

    user_code: str
    verification_uri: str
    expires_in_seconds: int


class MsalApplication(Protocol):
    """The slice of msal.PublicClientApplication this adapter uses.

    It exists so tests can drive sign-in and refresh without reaching
    login.microsoftonline.com, which is the external boundary being doubled.
    """

    def get_accounts(self) -> Sequence[object]: ...

    def acquire_token_silent(
        self, scopes: list[str], account: object
    ) -> Mapping[str, Any] | None: ...

    def initiate_device_flow(self, scopes: list[str]) -> Mapping[str, Any]: ...

    def acquire_token_by_device_flow(self, flow: Mapping[str, Any]) -> Mapping[str, Any]: ...


class DeviceCodeAuthenticator:
    """Issues Graph access tokens for a personal Microsoft account.

    Why sign-in is never attempted on demand: the device-code flow blocks for minutes
    waiting for a browser, and an MCP stdio server has nowhere to show the code, since
    stdout is the transport. Request paths call `get_access_token`, which only ever uses
    the cache and fails fast; `sign_in` is a deliberate out-of-band step.
    """

    def __init__(self, application: MsalApplication, token_cache: TokenCache) -> None:
        self._application = application
        self._token_cache = token_cache

    @classmethod
    def from_settings(cls, settings: Settings) -> "DeviceCodeAuthenticator":
        token_cache = KeyringTokenCache()
        application = PublicClientApplication(
            settings.client_id,
            authority=CONSUMERS_AUTHORITY,
            token_cache=token_cache.msal_cache,
        )
        return cls(application, token_cache)

    def get_access_token(self) -> str:
        account = self._cached_account()
        if account is None:
            raise NotAuthenticatedError(SIGN_IN_REQUIRED_MESSAGE)
        response = self._application.acquire_token_silent(list(SCOPES), account)
        self._token_cache.save_if_changed()
        if response is None:
            raise NotAuthenticatedError(SIGN_IN_REQUIRED_MESSAGE)
        return _read_access_token(response)

    def sign_in(self, show_prompt: Callable[[DeviceCodePrompt], None]) -> None:
        flow = self._application.initiate_device_flow(list(SCOPES))
        show_prompt(_read_device_code_prompt(flow))
        response = self._application.acquire_token_by_device_flow(flow)
        self._token_cache.save_if_changed()
        _read_access_token(response)

    def _cached_account(self) -> object | None:
        accounts = self._application.get_accounts()
        if not accounts:
            return None
        return accounts[0]


def _raise_for_error(response: Mapping[str, Any]) -> None:
    error = response.get("error")
    if error is None:
        return
    raise AuthenticationError(f"{error}: {response.get('error_description', '')}")


def _read_access_token(response: Mapping[str, Any]) -> str:
    _raise_for_error(response)
    access_token = response.get("access_token")
    if not isinstance(access_token, str):
        raise AuthenticationError("the token response carried no access token")
    return access_token


def _read_device_code_prompt(flow: Mapping[str, Any]) -> DeviceCodePrompt:
    _raise_for_error(flow)
    user_code = flow.get("user_code")
    verification_uri = flow.get("verification_uri")
    expires_in = flow.get("expires_in")
    if not (
        isinstance(user_code, str)
        and isinstance(verification_uri, str)
        and isinstance(expires_in, int)
    ):
        raise AuthenticationError("the device flow response was incomplete")
    return DeviceCodePrompt(
        user_code=user_code,
        verification_uri=verification_uri,
        expires_in_seconds=expires_in,
    )
