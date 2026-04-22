from pathlib import Path

from helpers.config.general_config import Settings


class ProjectPaths:
    """
    Minimal path resolver for slides pipeline artifacts.
    """

    def __init__(self, project_id: str, settings: Settings | None = None):
        self.project_id = str(project_id)
        self.settings = settings or Settings()
        self.base_path = Path(self.settings.Main_PATH)
        self.project_root = self.base_path / "data" / self.project_id

    def dir(self, key: str, prefer_existing: bool = False) -> Path:
        directories = {
            "responses": self.project_root / "responses",
            "json_files": self.project_root / "json_files",
            "pptx_files": self.project_root / "pptx_files",
            "content_files": self.project_root / "content_files",
            "media": self.project_root / "media",
            "images": self.project_root / "images",
        }
        if key not in directories:
            raise KeyError(f"Unknown project directory key: {key}")
        return directories[key]

