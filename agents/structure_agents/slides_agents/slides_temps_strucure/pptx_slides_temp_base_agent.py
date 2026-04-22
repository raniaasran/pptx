import json
import re
from typing import Any


class BaseSlidesStructureAgent:
    """
    Minimal compatibility implementation for the slides reconstruction agent.

    This keeps the same method contract expected by SlidesService and returns
    a valid slide JSON structure.
    """

    _SLIDE_HEADER_PATTERN = re.compile(r"(?mi)^slide\s+\d+\s*:")

    def __init__(self, settings: Any = None, generation_client: Any = None):
        self.settings = settings
        self.generation_client = generation_client

    def _split_topic_slides(self, topic_slides: str) -> list[str]:
        text = str(topic_slides or "").strip()
        if not text:
            return []

        headers = list(self._SLIDE_HEADER_PATTERN.finditer(text))
        if not headers:
            return [text]

        chunks: list[str] = []
        for index, header in enumerate(headers):
            start = header.end()
            end = headers[index + 1].start() if index + 1 < len(headers) else len(text)
            chunk = text[start:end].strip()
            chunks.append(chunk)
        return chunks

    def _next_text(self, lines: list[str], fallback: str) -> str:
        return lines.pop(0) if lines else fallback

    def reconstruct_slides(
        self,
        *,
        topic_slides: str,
        slides_dict: dict[str, Any],
        max_slides: int,
        content_slide_count: int,
    ) -> tuple[str, None, None]:
        template_keys = list(slides_dict.keys())[: max(1, int(max_slides or 1))]
        if not template_keys:
            raise ValueError("No templates available for slide reconstruction")

        chunks = self._split_topic_slides(topic_slides)
        expected_count = int(content_slide_count or len(chunks) or 1)

        if len(chunks) < expected_count:
            chunks.extend([""] * (expected_count - len(chunks)))
        elif len(chunks) > expected_count:
            chunks = chunks[:expected_count]

        payload: dict[str, dict[str, str]] = {}
        for index in range(expected_count):
            template_key = template_keys[index % len(template_keys)]
            placeholders = slides_dict[template_key]["placeholders"]
            raw_lines = [
                line.strip(" -\t")
                for line in chunks[index].splitlines()
                if line and line.strip()
            ]

            slide_data: dict[str, str] = {"template_key": template_key}
            for placeholder_key in placeholders.keys():
                lowered_key = str(placeholder_key).lower()
                if "picture" in lowered_key:
                    slide_data[str(placeholder_key)] = ""
                    continue

                fallback_value = f"Slide {index + 1}"
                if "title" in lowered_key:
                    fallback_value = f"Slide {index + 1} Title"
                slide_data[str(placeholder_key)] = self._next_text(raw_lines, fallback_value)

            payload[f"slide_{index + 1}"] = slide_data

        return json.dumps(payload, ensure_ascii=False, indent=2), None, None

