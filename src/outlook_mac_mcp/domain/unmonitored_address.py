"""Heuristic for addresses that probably nobody reads.

A reply to an automated sender is saved like any other draft; the risk is the user
believing it reached a person. So a match only adds a warning to the preview and never
blocks anything, which is why false positives are acceptable and the lists stay short.
"""

import re

from outlook_mac_mcp.domain.reply_recipients import address_key

UNMONITORED_LOCAL_PART = re.compile(
    r"(no[-_.]?reply|do[-_.]?not[-_.]?reply|mailer[-_.]?daemon|postmaster|bounces?"
    r"|notifications?|alerts?|automated|auto[-_.]?confirm)"
)
# A first label like this on a domain with at least three labels is how bulk-mail
# platforms usually name their sending hosts, e.g. someone@email.example.com.
BULK_SENDER_LABEL = re.compile(
    r"(e|em|email|mail|mailer|mailing|bounces?|notifications?|notify|news|newsletter|mkt"
    r"|marketing|mg|reply)\d*"
)
MIN_LABELS_FOR_SENDING_HOST = 3


def looks_unmonitored(address: str) -> bool:
    local_part, _, domain = address_key(address).rpartition("@")
    if not local_part or not domain:
        return False
    if UNMONITORED_LOCAL_PART.search(local_part):
        return True
    labels = domain.split(".")
    if len(labels) < MIN_LABELS_FOR_SENDING_HOST:
        return False
    return BULK_SENDER_LABEL.fullmatch(labels[0]) is not None
