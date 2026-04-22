from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from mini_layout_engine.engine.planning_engine import PlanningEngine


def _load_request(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        if not payload:
            raise ValueError(f"Request file is an empty list: {path}")
        first = payload[0]
        if not isinstance(first, dict):
            raise ValueError(f"First item must be an object in request list: {path}")
        return first
    if not isinstance(payload, dict):
        raise ValueError(f"Request file must contain an object or list of objects: {path}")
    return payload


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run mini layout planning for a JSON request and save output JSON."
    )
    parser.add_argument(
        "--request",
        default="mini_layout_engine/examples/requests/infographic_v1_demo_10_slides.json",
        help="Path to planning request JSON file",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Optional output path for plan result JSON",
    )
    args = parser.parse_args()

    request_path = Path(args.request).resolve()
    if not request_path.exists():
        raise FileNotFoundError(f"Request file not found: {request_path}")

    request_obj = _load_request(request_path)
    engine = PlanningEngine.from_default_specs()
    result = engine.plan_deck(request_obj)
    output_payload = result.to_dict()

    if args.output:
        output_path = Path(args.output).resolve()
    else:
        output_dir = Path("mini_layout_engine/examples/outputs").resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{request_path.stem}.plan_output.json"

    output_path.write_text(
        json.dumps(output_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "status": output_payload.get("status"),
                "slides_total": output_payload.get("summary", {}).get("slides_total", 0),
                "output_path": str(output_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
