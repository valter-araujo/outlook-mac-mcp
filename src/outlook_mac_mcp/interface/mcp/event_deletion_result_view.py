from pydantic import BaseModel, ConfigDict, Field


class EventDeletionResultView(BaseModel):
    """What delete_event hands back: confirmation of which event is gone."""

    model_config = ConfigDict(frozen=True)

    event_id: str = Field(description="Graph id of the event that was deleted.")
    deleted: bool = Field(default=True, description="Always true: the tool raises on failure.")

    @classmethod
    def from_event_id(cls, event_id: str) -> "EventDeletionResultView":
        return cls(event_id=event_id)
