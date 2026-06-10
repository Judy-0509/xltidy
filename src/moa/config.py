from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass
class Settings:
    base_url: str | None = None
    api_key: str | None = None
    model: str | None = None


def _env(name: str) -> str | None:
    # MOA_* is canonical; XLTIDY_* kept as a legacy fallback (pre-rename installs).
    return os.environ.get(f"MOA_{name}") or os.environ.get(f"XLTIDY_{name}")


def load_settings() -> Settings:
    return Settings(
        base_url=_env("QWEN_BASE_URL"),
        api_key=_env("QWEN_API_KEY"),
        model=_env("QWEN_MODEL"),
    )
