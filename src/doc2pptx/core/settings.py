# src/doc2pptx/core/settings.py
import os
import logging
from pathlib import Path
from typing import Dict, Optional
from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

# Get the logger for this module
logger = logging.getLogger(__name__)

class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        env_prefix="",
        extra="ignore"
    )

    openai_api_key: str
    unsplash_access_key: Optional[str] = None
    unsplash_secret_key: Optional[str] = None

    templates_dir: Path = Field(default_factory=lambda: Path("templates"))
    output_dir: Path = Field(default_factory=lambda: Path("output"))
    cache_dir: Path = Field(default_factory=lambda: Path("cache"))

    mermaid_cli_path: Optional[str] = Field(default="mmdc")

    openai_model: str = Field(default="gpt-4o")
    openai_temperature: float = Field(default=0.0)

    debug: bool = Field(default=False)
    layout_rules_path: Path = Field(default_factory=lambda: Path("layout/rules.yaml"))

    # custom_env_vars logic removed for simplicity in previous step

    @field_validator("templates_dir", "output_dir", "cache_dir", "layout_rules_path")
    @classmethod
    def validate_path_exists(cls, value: Path) -> Path:
        if not value.is_absolute():
            value = Path.cwd() / value
        if not value.exists():
            value.mkdir(parents=True, exist_ok=True)
        elif not value.is_dir():
            raise ValueError(f"Path {value} exists but is not a directory")
        return value

# Instantiate settings upon module import
try:
    settings = Settings()
    # Log the loaded API key (partially masked) for debugging
    # This log relies on the logger being configured by the entry point
    masked_key = f"{settings.openai_api_key[:5]}...{settings.openai_api_key[-4:]}" if settings.openai_api_key else "None"
    # Use debug level so it only shows up with -v or DEBUG log level
    logger.debug(f"Settings loaded. OpenAI API Key (masked): {masked_key}")
except Exception as e:
    logger.error(f"Failed to load settings: {e}")
    # Create a dummy settings object on failure to allow import without crash
    class DummySettings(BaseSettings): # Inherit BaseSettings to keep type compatibility
        openai_api_key: str = "" # Default to empty string on failure
        # Provide default values for other required fields
        templates_dir: Path = Path(".")
        output_dir: Path = Path(".")
        cache_dir: Path = Path(".")
        mermaid_cli_path: Optional[str] = None # Use None if no default
        openai_model: str = "gpt-4o"
        openai_temperature: float = 0.0
        debug: bool = False
        layout_rules_path: Path = Path(".")

    settings = DummySettings()
    # Log that dummy settings are being used
    logger.warning("Using dummy settings due to loading failure.")
