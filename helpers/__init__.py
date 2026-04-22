import json
from pathlib import Path

from helpers.config.general_config import Settings
from helpers.project_paths import ProjectPaths


def load_project_file(project_id: str, file_name: str, settings: Settings | None = None) -> dict:
    paths = ProjectPaths(project_id=project_id, settings=settings or Settings())
    safe_name = Path(str(file_name)).name
    if safe_name != str(file_name):
        raise ValueError("Invalid file name")

    candidates = [
        paths.dir("responses", prefer_existing=True) / safe_name,
        paths.project_root / safe_name,
    ]
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return json.loads(candidate.read_text(encoding="utf-8"))

    raise FileNotFoundError(
        "Project file was not found. Checked: " + ", ".join(str(path) for path in candidates)
    )

