from enum import StrEnum


class SearchScope(StrEnum):
    """Which part of an email a term is matched against.

    A closed set on purpose: the caller picks a scope, never writes one. The backend's
    own query syntax is built from this enum, so a term can never smuggle in a
    restriction of its own.
    """

    ANY = "any"
    SUBJECT = "subject"
    SENDER = "sender"
