import secrets
from collections.abc import Callable

from outlook_mac_mcp.domain.errors import DraftNotFoundError

TOKEN_BYTES = 16


class DraftStore[DraftType]:
    """Drafts of one kind, held in memory for the life of the server process.

    Not persisted on purpose: a draft is a pending confirmation, and a confirmation
    should not outlive the conversation that asked for it. A restart clears them all,
    and the client simply previews again.

    Tokens are unguessable rather than sequential so a call cannot reach a draft it was
    not given. Generic over the draft type so create, update and deletion each get their
    own store and their own token namespace: a token issued by one store is simply not a
    key in another, which is what makes a create token unusable for update or delete and
    vice versa, without any type-checking at the point of use.
    """

    def __init__(self) -> None:
        self._drafts: dict[str, DraftType] = {}

    def add(self, build: Callable[[str], DraftType]) -> DraftType:
        """`build` receives the freshly minted token and returns the draft to store
        under it, so the store never needs to know a specific draft type's fields.
        """
        token = secrets.token_urlsafe(TOKEN_BYTES)
        draft = build(token)
        self._drafts[token] = draft
        return draft

    def take(self, token: str) -> DraftType:
        """Return the draft and forget it, so a token is honoured at most once."""
        try:
            return self._drafts.pop(token)
        except KeyError:
            raise DraftNotFoundError("no pending draft for that token; preview again") from None
