from pydantic import BaseModel, Field


class SkillLoggingConfig(BaseModel):
    """Configuration for skill execution logging."""

    enabled: bool = Field(
        default=True,
        description="Whether skill usage and execution logging is enabled.",
    )
    log_to_file: bool = Field(
        default=True,
        description="Whether to write skill logs to JSONL files.",
    )
    log_dir: str | None = Field(
        default=None,
        description="Directory to store skill log files. If None, uses default logs/skills directory.",
    )
    max_days_to_keep: int = Field(
        default=30,
        description="Maximum number of days to keep skill log files before rotating.",
    )
