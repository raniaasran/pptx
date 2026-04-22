import json
import re
from pathlib import Path
from typing import Any

from agents.structure_agents.slides_agents.slides_temps_strucure.pptx_slides_temp_base_agent import (
    BaseSlidesStructureAgent,
)
from helpers import load_project_file
from helpers.config.general_config import Settings
from helpers.helper_functions.pptx_functions import resolve_media_reference_name
from helpers.project_paths import ProjectPaths
from persistence import get_persistence_facade
from persistence.enums import StoredFileRole
from persistence.utils import sanitize_filename, topic_slug_from_artifact_name
from services.content_services.errors import Phase2BadRequestError
from services.content_services.json_parsing import parse_json_payload
from services.content_services.slide_payloads import canonicalize_slide_payload
from services.content_services.slide_templates import map_dict, template_dict
from test_run.slides_automation import generate_pptx


DEFAULT_FRONTEND_TEMPLATE_ID = "MODERN"
DEFAULT_HELPER_TEMPLATE_FILE = "template_2.pptx"
FRONTEND_TEMPLATE_TO_HELPER_FILE = {
    "MINIMAL": DEFAULT_HELPER_TEMPLATE_FILE,
    "MODERN": DEFAULT_HELPER_TEMPLATE_FILE,
    "CORPORATE": DEFAULT_HELPER_TEMPLATE_FILE,
    "CREATIVE": DEFAULT_HELPER_TEMPLATE_FILE,
    "ELEGANT": DEFAULT_HELPER_TEMPLATE_FILE,
}


class SlidesService:
    _SLIDE_HEADER_PATTERN = re.compile(r"(?mi)^slide\s+\d+\s*:")

    def __init__(self, generation_client: Any | None = None, settings: Settings | None = None):
        self.settings = settings or Settings()
        self.generation_client = generation_client
        self.facade = get_persistence_facade(self.settings)

    def _frontend_draft_path(self, project_id: str) -> Path:
        paths = ProjectPaths(project_id=project_id, settings=self.settings)
        responses_dir = paths.dir("responses", prefer_existing=True)
        responses_dir.mkdir(parents=True, exist_ok=True)
        return responses_dir / "frontend_draft.json"

    def _load_frontend_draft(self, project_id: str) -> dict[str, Any]:
        draft_path = self._frontend_draft_path(project_id)
        if not draft_path.exists():
            return {}
        try:
            return json.loads(draft_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def get_selected_template_id(self, project_id: str) -> str:
        payload = self._load_frontend_draft(project_id)
        requirements = (
            payload.get("project", {}).get("requirements_json", {})
            if isinstance(payload, dict)
            else {}
        )
        output_configs = (
            requirements.get("output_configs", {}) if isinstance(requirements, dict) else {}
        )
        presentation_config = (
            output_configs.get("PRESENTATION", {})
            if isinstance(output_configs, dict)
            else {}
        )
        selected_template = str(
            presentation_config.get("presentation_template", DEFAULT_FRONTEND_TEMPLATE_ID)
        ).strip().upper()
        return (
            selected_template
            if selected_template in FRONTEND_TEMPLATE_TO_HELPER_FILE
            else DEFAULT_FRONTEND_TEMPLATE_ID
        )

    def _humanize_topic_slug(self, topic_slug: str) -> str:
        parts = [segment for segment in str(topic_slug).split("_") if segment]
        if not parts:
            return "Untitled Topic"
        return " ".join(parts).title()

    def resolve_template_file(self, template_id: str | None = None, project_id: str | None = None) -> Path:
        selected_template_id = str(template_id or "").strip().upper()
        if not selected_template_id and project_id:
            selected_template_id = self.get_selected_template_id(project_id)
        if selected_template_id not in FRONTEND_TEMPLATE_TO_HELPER_FILE:
            selected_template_id = DEFAULT_FRONTEND_TEMPLATE_ID

        helper_template_file = FRONTEND_TEMPLATE_TO_HELPER_FILE[selected_template_id]
        base_path = Path(self.settings.Main_PATH)
        candidate_paths = [
            base_path / "helpers" / "pptx_templates" / helper_template_file,
            base_path / "src" / "helpers" / "pptx_templates" / helper_template_file,
        ]

        for candidate in candidate_paths:
            if candidate.exists():
                return candidate

        raise ValueError("PPTX template file was not found")

    def list_generated_slides(self, project_id: str) -> dict[str, Any]:
        project_id = str(project_id)
        paths = ProjectPaths(project_id=project_id, settings=self.settings)
        pptx_dir = paths.dir("pptx_files", prefer_existing=True)
        pptx_dir.mkdir(parents=True, exist_ok=True)

        files = []
        for pptx_path in sorted(pptx_dir.glob("*.pptx")):
            topic_slug = topic_slug_from_artifact_name(pptx_path.name)
            files.append(
                {
                    "file_name": pptx_path.name,
                    "topic_slug": topic_slug,
                    "topic_title": self._humanize_topic_slug(topic_slug),
                }
            )

        template_id = self.get_selected_template_id(project_id)
        template_file_name = self.resolve_template_file(template_id=template_id).name
        return {
            "generated_slides": len(files),
            "generated_files": files,
            "template_id": template_id,
            "template_file_name": template_file_name,
        }

    def resolve_generated_slide_path(self, project_id: str, file_name: str) -> Path:
        safe_name = Path(str(file_name or "")).name
        if not safe_name or safe_name != str(file_name or ""):
            raise ValueError("Invalid slide file name")
        if Path(safe_name).suffix.lower() != ".pptx":
            raise ValueError("Only PPTX slide files are supported")

        paths = ProjectPaths(project_id=str(project_id), settings=self.settings)
        pptx_path = paths.dir("pptx_files", prefer_existing=True) / safe_name
        if not pptx_path.exists() or not pptx_path.is_file():
            raise ValueError("Slide file was not found")
        return pptx_path

    def _json_dir(self, project_id: str) -> Path:
        json_dir = ProjectPaths(project_id=project_id, settings=self.settings).dir(
            "json_files",
            prefer_existing=True,
        )
        json_dir.mkdir(parents=True, exist_ok=True)
        return json_dir

    def _pptx_dir(self, project_id: str) -> Path:
        pptx_dir = ProjectPaths(project_id=project_id, settings=self.settings).dir(
            "pptx_files",
            prefer_existing=True,
        )
        pptx_dir.mkdir(parents=True, exist_ok=True)
        return pptx_dir

    def _load_course_outline(self, project_id: str) -> list[dict[str, Any]]:
        syllabus = load_project_file(project_id=project_id, file_name="evaluation.json")
        course_outline = syllabus.get("course_outline")
        if not isinstance(course_outline, list):
            raise ValueError("Invalid syllabus payload: course_outline must be a list")
        return course_outline

    def _topic_response_name(self, topic_title: str) -> str:
        return f"{sanitize_filename(topic_title + '_r')}.json"

    def _topic_pptx_name(self, topic_title: str) -> str:
        return f"{sanitize_filename(topic_title + '_r')}.pptx"

    def _normalize_topic_titles(self, topic_titles: list[str] | None) -> tuple[list[str], set[str]]:
        cleaned_titles: list[str] = []
        normalized_titles: set[str] = set()
        for title in topic_titles or []:
            cleaned = str(title or "").strip()
            if not cleaned:
                continue
            normalized = cleaned.casefold()
            if normalized in normalized_titles:
                continue
            cleaned_titles.append(cleaned)
            normalized_titles.add(normalized)
        return cleaned_titles, normalized_titles

    def _topic_content_path(
        self,
        project_id: str,
        module_title: str,
        lesson_title: str,
        topic_title: str,
    ) -> Path:
        content_root = ProjectPaths(project_id=project_id, settings=self.settings).dir(
            "content_files",
            prefer_existing=True,
        )
        return (
            content_root
            / sanitize_filename(module_title)
            / sanitize_filename(lesson_title)
            / f"{sanitize_filename(topic_title)}.json"
        )

    def _load_topic_content(
        self,
        project_id: str,
        module_title: str,
        lesson_title: str,
        topic_title: str,
        topic_node: dict[str, Any] | None,
    ) -> tuple[str, dict[str, Any] | None]:
        if topic_node is not None:
            script_record = self.facade.get_topic_script(topic_node["id"])
            if isinstance(script_record, dict):
                markdown_body = script_record.get("markdown_body")
                if isinstance(markdown_body, str) and markdown_body.strip():
                    return markdown_body.strip(), script_record

        topic_path = self._topic_content_path(project_id, module_title, lesson_title, topic_title)
        if not topic_path.exists() or not topic_path.is_file():
            raise ValueError(f"Topic content was not found for topic: {topic_title}")

        try:
            payload = json.loads(topic_path.read_text(encoding="utf-8"))
        except Exception as exc:
            raise ValueError(f"Invalid topic content payload for topic: {topic_title}") from exc

        topic_content = payload.get("content") if isinstance(payload, dict) else None
        if not isinstance(topic_content, str) or not topic_content.strip():
            raise ValueError(f"Topic content was not found for topic: {topic_title}")

        return topic_content.strip(), None

    def _count_expected_slides(self, topic_content: str) -> int:
        if not isinstance(topic_content, str) or not topic_content.strip():
            return 0
        return len(self._SLIDE_HEADER_PATTERN.findall(topic_content))

    def _canonicalize_slide_payload(
        self,
        *,
        topic_title: str,
        payload: dict[str, Any],
        expected_slide_count: int,
    ) -> dict[str, Any]:
        structured_slides = canonicalize_slide_payload(
            payload,
            available_templates=template_dict,
        )
        if expected_slide_count and len(structured_slides) != expected_slide_count:
            raise ValueError(
                f"Expected {expected_slide_count} slides for topic '{topic_title}' "
                f"but generated {len(structured_slides)}"
            )
        return structured_slides

    def _repair_slide_payload(
        self,
        topic_title: str,
        raw_answer: str,
        *,
        topic_content: str,
        expected_slide_count: int,
    ) -> dict[str, Any]:
        if self.generation_client is None:
            raise Phase2BadRequestError(
                "SLIDE_JSON_INVALID",
                f"Slide JSON generation failed for topic '{topic_title}'.",
            )

        repair_history = [
            self.generation_client.construct_prompt(
                prompt=(
                    "You repair malformed slide-structure responses. "
                    "Return ONLY valid JSON. No markdown, no code fences, no commentary."
                ),
                role=self.generation_client.enums.SYSTEM.value,
            )
        ]
        repair_prompt = "\n".join(
            [
                "Repair the following slide-structure response into a valid JSON object.",
                "Return ONLY JSON. No markdown. No commentary.",
                "",
                "STRICT OUTPUT RULES:",
                f"- Return exactly {expected_slide_count or 'the required number of'} top-level slide objects.",
                '- Top-level keys must be sequential: "slide_1", "slide_2", ...',
                '- Each slide object must include a "template_key" field.',
                "- template_key must be one of the keys present in TEMPLATE_DICT below.",
                "- Each template_key may be used at most once.",
                "- Preserve the original slide order from TOPIC_SLIDES.",
                "- Do not merge or skip slides.",
                "- Keep all existing shape fields for each chosen template.",
                "",
                "TEMPLATE_DICT:",
                json.dumps(template_dict, ensure_ascii=False, indent=2),
                "",
                "TOPIC_SLIDES:",
                topic_content,
                "",
                "RAW_RESPONSE_TO_REPAIR:",
                "",
                raw_answer,
            ]
        )
        repaired_answer = self.generation_client.generate_text(
            prompt=repair_prompt,
            chat_history=repair_history,
        )
        try:
            return parse_json_payload(repaired_answer)
        except Exception as exc:
            raise Phase2BadRequestError(
                "SLIDE_JSON_INVALID",
                f"Slide JSON generation failed for topic '{topic_title}': {exc}",
            ) from exc

    def _generate_slide_payload(self, topic_title: str, topic_content: str) -> dict[str, Any]:
        if self.generation_client is None:
            raise ValueError("Slides generation client is required to build slide JSON files")

        expected_slide_count = self._count_expected_slides(topic_content)
        restructure_agent = BaseSlidesStructureAgent(
            settings=self.settings,
            generation_client=self.generation_client,
        )
        answer, _, _ = restructure_agent.reconstruct_slides(
            topic_slides=topic_content,
            slides_dict=template_dict,
            max_slides=len(template_dict),
            content_slide_count=expected_slide_count,
        )
        try:
            structured_slides = self._canonicalize_slide_payload(
                topic_title=topic_title,
                payload=parse_json_payload(answer),
                expected_slide_count=expected_slide_count,
            )
        except Exception:
            repaired_payload = self._repair_slide_payload(
                topic_title=topic_title,
                raw_answer=str(answer or ""),
                topic_content=topic_content,
                expected_slide_count=expected_slide_count,
            )
            structured_slides = self._canonicalize_slide_payload(
                topic_title=topic_title,
                payload=repaired_payload,
                expected_slide_count=expected_slide_count,
            )
        if not structured_slides:
            raise ValueError(f"No slide structure generated for topic: {topic_title}")
        if not isinstance(structured_slides, dict):
            raise ValueError(f"Invalid slide structure generated for topic: {topic_title}")
        return structured_slides

    def _normalize_media_placeholders(
        self,
        project_id: str,
        structured_slides: dict[str, Any],
    ) -> dict[str, Any]:
        if not isinstance(structured_slides, dict):
            return structured_slides

        for slide_payload in structured_slides.values():
            if not isinstance(slide_payload, dict):
                continue
            for key, value in list(slide_payload.items()):
                if "picture" not in str(key).lower():
                    continue
                if not isinstance(value, str) or not value.strip():
                    continue
                context_text = " ".join(
                    str(item).strip()
                    for item_key, item in slide_payload.items()
                    if str(item_key) != "template_key"
                    if "picture" not in str(item_key).lower()
                    and isinstance(item, str)
                    and str(item).strip()
                )
                try:
                    slide_payload[key] = resolve_media_reference_name(
                        project_id,
                        value,
                        settings=self.settings,
                        context_text=context_text,
                    )
                except Exception:
                    # Do not keep invented or unresolvable media names in the saved
                    # slide payload, otherwise PPTX generation fails later.
                    slide_payload[key] = ""
        return structured_slides

    def _write_slide_json(
        self,
        project_id: str,
        topic_title: str,
        topic_node: dict[str, Any] | None,
        script_record: dict[str, Any] | None,
        structured_slides: dict[str, Any],
    ) -> Path:
        json_path = self._json_dir(project_id) / self._topic_response_name(topic_title)
        json_path.write_text(
            json.dumps(structured_slides, ensure_ascii=False, indent=4),
            encoding="utf-8",
        )

        if topic_node is not None:
            slide_json_file_record = self.facade.register_local_file(
                project_id=int(project_id),
                role=StoredFileRole.SLIDE_JSON,
                path=str(json_path),
                node_id=topic_node["id"],
                metadata={"content": structured_slides, "topic_title": topic_title},
            )
            self.facade.upsert_slide_deck(
                topic_node_id=topic_node["id"],
                slides_json=structured_slides,
                script_id=script_record["id"] if isinstance(script_record, dict) else None,
                slide_json_file_id=slide_json_file_record["id"] if slide_json_file_record else None,
            )

        return json_path

    def _rebuild_slide_json_files(self, project_id: str, topic_titles: list[str] | None = None) -> list[Path]:
        project_id = str(project_id)
        json_dir = self._json_dir(project_id)
        course_outline = self._load_course_outline(project_id)
        requested_titles, requested_title_keys = self._normalize_topic_titles(topic_titles)
        expected_json_names: set[str] = set()
        expected_pptx_names: set[str] = set()
        generated_files: list[Path] = []
        found_topic_keys: set[str] = set()

        for learning_module in course_outline:
            module_title = str(learning_module.get("module_title") or "").strip()
            for lesson in learning_module.get("lessons", []):
                lesson_title = str(lesson.get("lesson_title") or "").strip()
                for topic in lesson.get("topics", []):
                    topic_title = str(topic.get("topic_title") or "").strip()
                    if not topic_title:
                        raise ValueError("Topic is missing topic_title")
                    topic_key = topic_title.casefold()
                    if requested_title_keys and topic_key not in requested_title_keys:
                        continue
                    found_topic_keys.add(topic_key)

                    expected_json_names.add(self._topic_response_name(topic_title))
                    expected_pptx_names.add(self._topic_pptx_name(topic_title))

                    topic_node = self.facade.resolve_topic_node(
                        project_id=int(project_id),
                        topic_title=topic_title,
                    )
                    topic_content, script_record = self._load_topic_content(
                        project_id=project_id,
                        module_title=module_title,
                        lesson_title=lesson_title,
                        topic_title=topic_title,
                        topic_node=topic_node,
                    )
                    structured_slides = self._generate_slide_payload(
                        topic_title=topic_title,
                        topic_content=topic_content,
                    )
                    structured_slides = self._normalize_media_placeholders(
                        project_id=project_id,
                        structured_slides=structured_slides,
                    )
                    generated_files.append(
                        self._write_slide_json(
                            project_id=project_id,
                            topic_title=topic_title,
                            topic_node=topic_node,
                            script_record=script_record,
                            structured_slides=structured_slides,
                        )
                    )

        if requested_title_keys:
            missing_titles = [title for title in requested_titles if title.casefold() not in found_topic_keys]
            if missing_titles:
                raise Phase2BadRequestError(
                    "TOPIC_NOT_FOUND",
                    f"Selected topic(s) were not found in the confirmed outline: {', '.join(missing_titles)}",
                )
        else:
            for stale_json_path in json_dir.glob("*.json"):
                if stale_json_path.name not in expected_json_names:
                    stale_json_path.unlink(missing_ok=True)

            pptx_dir = self._pptx_dir(project_id)
            for stale_pptx_path in pptx_dir.glob("*.pptx"):
                if stale_pptx_path.name not in expected_pptx_names:
                    stale_pptx_path.unlink(missing_ok=True)

        return sorted(generated_files)

    def generate_slides(self, project_id: str, topic_titles: list[str] | None = None) -> dict[str, Any]:
        project_id = str(project_id)
        json_dir = self._json_dir(project_id)
        requested_titles, requested_title_keys = self._normalize_topic_titles(topic_titles)
        if self.generation_client is None:
            json_files = sorted(json_dir.glob("*.json"))
            if requested_title_keys:
                json_files = [
                    path
                    for path in json_files
                    if topic_slug_from_artifact_name(path.name).replace("_", " ").casefold() in requested_title_keys
                    or path.stem.removesuffix("_r").replace("_", " ").casefold() in requested_title_keys
                ]
        else:
            json_files = self._rebuild_slide_json_files(project_id, topic_titles=requested_titles)
        if not json_files and self.facade.database_enabled:
            for deck in self.facade.list_slide_decks(project_id=int(project_id)):
                file_name = f"{deck['topic_slug']}_r.json"
                target_path = json_dir / file_name
                target_path.write_text(
                    json.dumps(deck["slides_json"], ensure_ascii=False, indent=4),
                    encoding="utf-8",
                )
            json_files = sorted(json_dir.glob("*.json"))
            if requested_title_keys:
                json_files = [
                    path
                    for path in json_files
                    if topic_slug_from_artifact_name(path.name).replace("_", " ").casefold() in requested_title_keys
                    or path.stem.removesuffix("_r").replace("_", " ").casefold() in requested_title_keys
                ]
        if not json_files:
            raise Phase2BadRequestError(
                "SLIDES_NOT_FOUND",
                "No generated slide json files were found for the selected topic(s).",
            )

        template_id = self.get_selected_template_id(project_id)
        pptx_template_path = str(self.resolve_template_file(template_id=template_id))

        generate_pptx(
            project_id=project_id,
            template_dict=template_dict,
            map_dict=map_dict,
            pptx_template_path=pptx_template_path,
            file_names=[path.name for path in json_files],
        )

        if self.facade.database_enabled:
            pptx_dir = self._pptx_dir(project_id)
            expected_pptx_names = {self._topic_pptx_name(path.stem.removesuffix("_r").replace("_", " ")) for path in json_files}
            for pptx_path in sorted(pptx_dir.glob("*.pptx")):
                if expected_pptx_names and pptx_path.name not in expected_pptx_names:
                    continue
                topic_slug = topic_slug_from_artifact_name(pptx_path.name)
                topic_node = self.facade.resolve_topic_node(
                    project_id=int(project_id),
                    slug=topic_slug,
                )
                if topic_node is None:
                    continue
                file_record = self.facade.register_local_file(
                    project_id=int(project_id),
                    role=StoredFileRole.SLIDE_PPTX,
                    path=str(pptx_path),
                    node_id=topic_node["id"],
                    metadata={"topic_slug": topic_slug},
                )
                deck_record = self.facade.get_slide_deck(topic_node["id"])
                if deck_record:
                    self.facade.upsert_slide_deck(
                        topic_node_id=topic_node["id"],
                        slides_json=deck_record["slides_json"],
                        slide_json_file_id=deck_record["slide_json_file_id"],
                        pptx_file_id=file_record["id"] if file_record else None,
                    )

        result = self.list_generated_slides(project_id)
        if requested_title_keys:
            result["generated_files"] = [
                file
                for file in result.get("generated_files", [])
                if str(file.get("topic_title") or "").strip().casefold() in requested_title_keys
            ]
        result["generated_slides"] = len(json_files)
        result["template_id"] = template_id
        return result
