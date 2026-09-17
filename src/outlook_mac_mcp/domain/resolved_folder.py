from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ResolvedFolder:
    """What a mail tool's raw `folder` argument resolved to.

    `folder_ids` are the Graph folder identifiers to actually query -- well-known names
    for a well-known selection or `all`, or the one Graph id a custom path resolved to.
    `echo` is what the tool's output should show back as `folder`: the well-known/`all`
    description, or the custom path exactly as given, never a resolved Graph id.
    """

    folder_ids: tuple[str, ...]
    echo: str
