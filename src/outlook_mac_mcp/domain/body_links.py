import re

LINK_PATTERN = re.compile(r"(?:https?://|www\.)[^\s<>\"']+", re.IGNORECASE)
TRAILING_PUNCTUATION = ".,;:!?)]}'\""


def find_links(text: str) -> tuple[str, ...]:
    """Every web address in a plain-text body, in order, without repeats.

    Escaping stops markup from becoming a live tag, but a bare URL stays a working link
    and Outlook may make it clickable. A URL can carry data in its query string to the
    site it names, so the preview shows each one for the user to check. Sentence
    punctuation right after a URL is not part of it.
    """
    links: dict[str, None] = {}
    for match in LINK_PATTERN.finditer(text):
        link = match.group().rstrip(TRAILING_PUNCTUATION)
        if link.casefold() not in ("www", "http://", "https://"):
            links.setdefault(link, None)
    return tuple(links)
