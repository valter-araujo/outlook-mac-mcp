from pydantic import BaseModel, ConfigDict, Field


class EmailCountView(BaseModel):
    model_config = ConfigDict(frozen=True)

    total: int = Field(description="Exactly how many emails match the filters.")
