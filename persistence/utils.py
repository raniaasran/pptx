import re
from pathlib import Path


def sanitize_filename(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return "untitled"
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-z0-9._-]", "_", text)
    text = re.sub(r"_+", "_", text).strip("._")
    return text or "untitled"


def topic_slug_from_artifact_name(file_name: str) -> str:
    stem = Path(str(file_name or "")).stem
    if stem.lower().endswith("_r"):
        stem = stem[:-2]
    return sanitize_filename(stem)

