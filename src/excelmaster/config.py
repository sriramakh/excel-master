"""Configuration settings for Excel Master."""
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # LLM providers
    openai_api_key: str = ""
    minimax_api_token: str = ""
    llm_provider: str = "openai"          # "openai" | "minimax"
    model: str = "gpt-4o"
    minimax_model: str = "MiniMax-Text-01"
    max_tokens: int = 4096
    temperature: float = 0.3
    max_retries: int = 3

    # Paths
    data_dir: Path = Path("data")
    output_dir: Path = Path("output")

    @property
    def active_model(self) -> str:
        return self.minimax_model if self.llm_provider == "minimax" else self.model

    @property
    def active_api_key(self) -> str:
        return self.minimax_api_token if self.llm_provider == "minimax" else self.openai_api_key


_settings: Settings | None = None


def get_settings() -> Settings:
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings
