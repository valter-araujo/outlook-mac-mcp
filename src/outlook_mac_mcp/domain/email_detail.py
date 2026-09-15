from dataclasses import dataclass

from outlook_mac_mcp.domain.email import Email


@dataclass(frozen=True, slots=True)
class EmailDetail:
    """An email together with its body.

    Kept apart from Email because listing never fetches bodies: Graph would return the
    full text of every message in a page, and the list tools deliberately carry only a
    preview. Composing rather than extending keeps one definition of what an email is.

    `body` is third-party text. Nothing in this project interprets it.
    """

    email: Email
    body: str
