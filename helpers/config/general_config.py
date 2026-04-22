import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Settings:
    """
    Minimal settings object expected by the slides pipeline.
    """

    Main_PATH: str = field(
        default_factory=lambda: os.getenv(
            "MAIN_PATH",
            str(Path(__file__).resolve().parents[2]),
        )
    )

