from outlook_mac_mcp.domain.body_links import find_links


def link_warnings(body: str) -> tuple[str, ...]:
    """One warning per web address in the body, so the user checks each before saving.

    Never blocks: links in a reply are normal. The risk is a link the user did not
    write, such as one copied from untrusted email content, carrying data in its query
    string to the site it names.
    """
    return tuple(
        f"the body links to {link}; check it is a link you meant to include, since a "
        "link can carry data to the site it points to"
        for link in find_links(body)
    )
