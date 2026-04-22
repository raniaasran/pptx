import json
from pathlib import Path
from typing import Iterable

from PptxHelperClass import PPTXHelper
from helpers.config.general_config import Settings
from helpers.project_paths import ProjectPaths
from pptx_functions import (
    compute_reorder_list,
    fill_presentation,
    keep_slides_by_index,
    reorder_slides,
)


def _selected_json_paths(json_dir: Path, file_names: Iterable[str] | None) -> list[Path]:
    if file_names:
        return [json_dir / Path(str(name)).name for name in file_names]
    return sorted(json_dir.glob("*.json"))


def generate_pptx(
    *,
    project_id: str,
    template_dict: dict,
    map_dict: dict,
    pptx_template_path: str,
    file_names: list[str] | None = None,
) -> list[Path]:
    settings = Settings()
    paths = ProjectPaths(project_id=project_id, settings=settings)
    json_dir = paths.dir("json_files", prefer_existing=True)
    pptx_dir = paths.dir("pptx_files", prefer_existing=True)
    pptx_dir.mkdir(parents=True, exist_ok=True)

    json_paths = _selected_json_paths(json_dir=json_dir, file_names=file_names)
    generated_paths: list[Path] = []

    for json_path in json_paths:
        if not json_path.exists() or not json_path.is_file():
            raise FileNotFoundError(f"Slide JSON file was not found: {json_path}")

        response_payload = json.loads(json_path.read_text(encoding="utf-8"))
        helper = PPTXHelper(str(pptx_template_path))

        updated_template, agent_selected = keep_slides_by_index(
            helper=helper,
            response=response_payload,
            template=template_dict,
            map_dict=map_dict,
        )
        fill_presentation(
            project_id=project_id,
            helper=helper,
            template_dict=updated_template,
            content_dict=response_payload,
        )
        reorder_slides(helper, compute_reorder_list(agent_selected))

        output_path = pptx_dir / f"{json_path.stem}.pptx"
        helper.save(str(output_path))
        generated_paths.append(output_path)

    return generated_paths

