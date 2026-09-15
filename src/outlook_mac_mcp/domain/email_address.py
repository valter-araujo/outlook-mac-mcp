from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class EmailAddress:
    address: str
    display_name: str = ""
