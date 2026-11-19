from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class CreatedMailDraft:
    """The saved draft as the mailbox reports it: just the handle to find it again."""

    id: str
