from enum import StrEnum


class ShowAs(StrEnum):
    """Graph's event.showAs (free/busy status) values, used verbatim so the adapter
    needs no mapping.
    """

    FREE = "free"
    TENTATIVE = "tentative"
    BUSY = "busy"
    OOF = "oof"
    WORKING_ELSEWHERE = "workingElsewhere"
