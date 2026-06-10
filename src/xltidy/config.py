from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass
class Settings:
    base_url: str | None = None
    api_key: str | None = None
    model: str | None = None


def load_settings() -> Settings:
    return Settings(
        base_url=os.environ.get("XLTIDY_QWEN_BASE_URL"),
        api_key=os.environ.get("XLTIDY_QWEN_API_KEY"),
        model=os.environ.get("XLTIDY_QWEN_MODEL"),
    )
