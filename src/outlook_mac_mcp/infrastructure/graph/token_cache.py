from typing import Protocol

import keyring
from keyring.errors import KeyringError
from msal import SerializableTokenCache

from outlook_mac_mcp.infrastructure.graph.errors import TokenCacheError

KEYCHAIN_SERVICE = "outlook-mac-mcp"
KEYCHAIN_ACCOUNT = "msal-token-cache"


class TokenCache(Protocol):
    """What the authenticator needs from a cache, so tests can supply a fake."""

    def save_if_changed(self) -> None: ...


class KeyringTokenCache:
    """An MSAL token cache persisted as a single Keychain generic password.

    Why the Keychain and not a file: the cache holds a refresh token, a long-lived
    credential. Keychain items are protected by the login password and per-application
    ACLs; a file under the home directory is readable by anything the user runs.

    Concurrent writers are not handled. Two server processes signed in to the same
    account would overwrite each other's state, and the loser signs in again. An MCP
    stdio server is started one per client, so this costs a sign-in, never a token leak.
    """

    def __init__(self, service: str = KEYCHAIN_SERVICE, account: str = KEYCHAIN_ACCOUNT) -> None:
        self._service = service
        self._account = account
        self._cache = SerializableTokenCache()
        self._load()

    @property
    def msal_cache(self) -> SerializableTokenCache:
        return self._cache

    def save_if_changed(self) -> None:
        if not self._cache.has_state_changed:
            return
        try:
            keyring.set_password(self._service, self._account, self._cache.serialize())
        except KeyringError as error:
            raise TokenCacheError("could not write the token cache to the Keychain") from error

    def _load(self) -> None:
        try:
            state = keyring.get_password(self._service, self._account)
        except KeyringError as error:
            raise TokenCacheError("could not read the token cache from the Keychain") from error
        if state is None:
            return
        self._cache.deserialize(state)
