from collections.abc import Mapping
from typing import Any, Protocol

import httpx

from outlook_mac_mcp.infrastructure.graph.errors import (
    GraphRequestError,
    GraphResponseError,
    UnsupportedHostError,
)

GRAPH_HOST = "graph.microsoft.com"
GRAPH_BASE_URL = f"https://{GRAPH_HOST}/v1.0"
DEFAULT_TIMEOUT_SECONDS = 30.0
UNKNOWN_ERROR_CODE = "no error code"
UNKNOWN_REQUEST_ID = "unknown"


class AccessTokenProvider(Protocol):
    """What the client needs from the authenticator, so tests can supply a fake."""

    def get_access_token(self) -> str: ...


class GraphClient:
    """HTTP access to Microsoft Graph, and the only place in the project that reaches it.

    Every request is pinned to graph.microsoft.com and redirects are never followed, so a
    redirect or a caller passing an absolute URL cannot replay the bearer token to another
    host. The token is attached only after the target host has been verified.
    """

    def __init__(
        self,
        token_provider: AccessTokenProvider,
        http_client: httpx.Client | None = None,
    ) -> None:
        self._token_provider = token_provider
        self._http_client = http_client if http_client is not None else _build_http_client()

    def get(self, path: str, parameters: Mapping[str, str | int]) -> Mapping[str, Any]:
        request = self._http_client.build_request("GET", _relative_path(path), params=parameters)
        _reject_foreign_host(request.url)
        request.headers["Authorization"] = f"Bearer {self._token_provider.get_access_token()}"
        return _read_payload(self._http_client.send(request))

    def close(self) -> None:
        self._http_client.close()


def _build_http_client() -> httpx.Client:
    return httpx.Client(
        base_url=GRAPH_BASE_URL,
        timeout=DEFAULT_TIMEOUT_SECONDS,
        follow_redirects=False,
    )


def _relative_path(path: str) -> str:
    """Reject anything that could carry its own host, including protocol-relative paths."""
    if not path.startswith("/") or path.startswith("//"):
        raise UnsupportedHostError(f"paths must be relative to {GRAPH_BASE_URL}")
    return path


def _reject_foreign_host(url: httpx.URL) -> None:
    if url.scheme != "https" or url.host != GRAPH_HOST:
        raise UnsupportedHostError(f"refusing to send a Graph request to {url.scheme}://{url.host}")


def _read_payload(response: httpx.Response) -> Mapping[str, Any]:
    if not response.is_success:
        raise GraphRequestError(_describe_failure(response))
    payload = _decode_json(response)
    if not isinstance(payload, dict):
        raise GraphResponseError("Graph returned a JSON value that is not an object")
    return payload


def _describe_failure(response: httpx.Response) -> str:
    """Report the status, Graph's error code and the request id, and nothing from the body.

    A Graph error body echoes the query and can carry mailbox content, so only these three
    fields are safe to surface or log.
    """
    request_id = response.headers.get("request-id", UNKNOWN_REQUEST_ID)
    return (
        f"Graph returned {response.status_code} ({_error_code(response)}); request-id {request_id}"
    )


def _error_code(response: httpx.Response) -> str:
    """Best effort on purpose: an unparsable error body must not mask the status code."""
    try:
        payload = response.json()
    except ValueError:
        return UNKNOWN_ERROR_CODE
    if not isinstance(payload, dict):
        return UNKNOWN_ERROR_CODE
    error = payload.get("error")
    if not isinstance(error, dict):
        return UNKNOWN_ERROR_CODE
    code = error.get("code")
    return code if isinstance(code, str) else UNKNOWN_ERROR_CODE


def _decode_json(response: httpx.Response) -> object:
    try:
        return response.json()
    except ValueError as error:
        raise GraphResponseError("Graph returned a body that is not JSON") from error
