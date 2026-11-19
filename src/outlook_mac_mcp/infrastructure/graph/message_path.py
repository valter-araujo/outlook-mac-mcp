from http import HTTPStatus
from typing import NoReturn
from urllib.parse import quote

from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from outlook_mac_mcp.infrastructure.graph.mail_repository import MALFORMED_ID_ERROR_CODE

MESSAGES_PATH = "/me/messages"


def message_path(email_id: str) -> str:
    """Percent-encoded for the same reason GraphMailRepository.get_by_id encodes it: the
    id is caller-supplied and could otherwise change the path.
    """
    return f"{MESSAGES_PATH}/{quote(email_id, safe='')}"


def raise_message_lookup_error(error: GraphRequestError, email_id: str) -> NoReturn:
    """Translate the two expected failures of addressing a message by id; anything else
    is an infrastructure failure and goes through unchanged.
    """
    if error.error_code == MALFORMED_ID_ERROR_CODE:
        raise InvalidRequestError(f"malformed email id: {email_id}") from error
    if error.status_code == HTTPStatus.NOT_FOUND:
        raise EmailNotFoundError(f"no email with id {email_id}") from error
    raise error
