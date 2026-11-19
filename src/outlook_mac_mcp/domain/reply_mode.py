from enum import StrEnum


class ReplyMode(StrEnum):
    """Which of Graph's two reply actions builds the draft.

    The values double as the tool argument, so a caller writes exactly what it means.
    """

    REPLY = "reply"
    REPLY_ALL = "reply_all"
