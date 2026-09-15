import json
from dataclasses import dataclass, field

import keyring
import pytest
from keyring.errors import KeyringError

from outlook_mac_mcp.infrastructure.graph.errors import TokenCacheError
from outlook_mac_mcp.infrastructure.graph.token_cache import KeyringTokenCache

STORED_STATE = '{"AccessToken": {}, "RefreshToken": {}, "Account": {}}'


@dataclass
class KeychainSpy:
    """Stands in for the macOS Keychain, the external boundary of this class."""

    stored: str | None = None
    writes: list[str] = field(default_factory=list)

    def get_password(self, service: str, account: str) -> str | None:
        return self.stored

    def set_password(self, service: str, account: str, password: str) -> None:
        self.writes.append(password)


@pytest.fixture
def keychain(monkeypatch: pytest.MonkeyPatch) -> KeychainSpy:
    spy = KeychainSpy()
    monkeypatch.setattr(keyring, "get_password", spy.get_password)
    monkeypatch.setattr(keyring, "set_password", spy.set_password)
    return spy


def test_loads_state_previously_stored_in_the_keychain(keychain: KeychainSpy) -> None:
    keychain.stored = STORED_STATE

    cache = KeyringTokenCache()

    assert json.loads(cache.msal_cache.serialize()) == json.loads(STORED_STATE)


def test_starts_empty_when_the_keychain_holds_nothing(keychain: KeychainSpy) -> None:
    cache = KeyringTokenCache()

    assert json.loads(cache.msal_cache.serialize()) == {}


def test_does_not_write_when_the_cache_is_unchanged(keychain: KeychainSpy) -> None:
    cache = KeyringTokenCache()

    cache.save_if_changed()

    assert keychain.writes == []


def test_writes_the_serialized_cache_when_it_changed(keychain: KeychainSpy) -> None:
    keychain.stored = STORED_STATE
    cache = KeyringTokenCache()
    cache.msal_cache.has_state_changed = True

    cache.save_if_changed()

    assert [json.loads(written) for written in keychain.writes] == [json.loads(STORED_STATE)]


def test_raises_when_the_keychain_cannot_be_read(monkeypatch: pytest.MonkeyPatch) -> None:
    def refuse(service: str, account: str) -> str | None:
        raise KeyringError("no backend")

    monkeypatch.setattr(keyring, "get_password", refuse)

    with pytest.raises(TokenCacheError):
        KeyringTokenCache()


def test_raises_when_the_keychain_cannot_be_written(
    keychain: KeychainSpy, monkeypatch: pytest.MonkeyPatch
) -> None:
    cache = KeyringTokenCache()
    cache.msal_cache.has_state_changed = True

    def refuse(service: str, account: str, password: str) -> None:
        raise KeyringError("locked")

    monkeypatch.setattr(keyring, "set_password", refuse)

    with pytest.raises(TokenCacheError):
        cache.save_if_changed()
