def totals_guidance(narrowing_advice: str) -> str:
    """The one wording every listing tool uses to make the client report its totals.

    A model that sees a full page and no total tends to treat it as the whole folder,
    so the description spells out both the sentence to say and what to suggest next.
    """
    return (
        "The output carries returned, total and total_is_exact. Always tell the user "
        '"showing N of M", where N is returned and M is total, or "showing N of at least '
        'M" when total_is_exact is false. When M exceeds N, say that more exist than '
        f"were shown and suggest narrowing the scope: {narrowing_advice}. Never conclude "
        "that something is absent from a page that does not hold every match."
    )
