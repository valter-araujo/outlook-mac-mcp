from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.sender_count import SenderCount
from outlook_mac_mcp.domain.sender_ranking import SenderRanking


class SenderCountView(BaseModel):
    model_config = ConfigDict(frozen=True)

    address: str
    name: str
    count: int = Field(description="How many scanned emails this sender sent.")

    @classmethod
    def from_count(cls, item: SenderCount) -> "SenderCountView":
        return cls(address=item.sender.address, name=item.sender.display_name, count=item.count)


class SenderRankingView(BaseModel):
    model_config = ConfigDict(frozen=True)

    senders: list[SenderCountView] = Field(description="Senders, most emails first.")
    scanned: int = Field(description="How many emails were counted to build the ranking.")
    total: int = Field(description="Exactly how many emails match the filters.")
    coverage_is_complete: bool = Field(
        description=(
            "Whether every matching email was scanned. When false, scanned is less than "
            "total and the ranking covers only the scanned emails."
        )
    )

    @classmethod
    def from_ranking(cls, ranking: SenderRanking) -> "SenderRankingView":
        return cls(
            senders=[SenderCountView.from_count(item) for item in ranking.senders],
            scanned=ranking.scanned,
            total=ranking.total,
            coverage_is_complete=ranking.coverage_is_complete,
        )
