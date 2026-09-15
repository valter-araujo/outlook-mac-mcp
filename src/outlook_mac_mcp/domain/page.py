from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Page[Item]:
    """One page of a listing, together with how much the listing holds in all.

    A page alone cannot tell a caller whether it saw everything: a full page may be the
    end of the list or the start of hundreds more. `total` answers that, and
    `total_is_exact` says whether the backend counted everything or gave up at a bound,
    in which case `total` is a lower bound and never an overstatement.
    """

    items: tuple[Item, ...]
    total: int
    total_is_exact: bool
