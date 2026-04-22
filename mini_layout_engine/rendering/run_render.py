from __future__ import annotations

import argparse
import json
from pathlib import Path

from mini_layout_engine.rendering.renderer import (
    MiniPptxRenderer,
    load_json_payload,
    normalize_request_payload,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render a PPTX from mini layout engine request or plan payload."
    )
    parser.add_argument(
        "--request",
        default="mini_layout_engine/examples/requests/infographic_v1_demo_10_slides.json",
        help="Path to request JSON (object or list with one object).",
    )
    parser.add_argument(
        "--plan",
        default="",
        help="Optional plan output JSON path. If provided, this is used directly.",
    )
    parser.add_argument(
        "--output",
        default="mini_layout_engine/examples/outputs/infographic_v1_demo_10_slides.pptx",
        help="Output PPTX path.",
    )
    args = parser.parse_args()

    renderer = MiniPptxRenderer()
    output_path = Path(args.output).resolve()

    if args.plan:
        plan_payload = load_json_payload(args.plan)
        result = renderer.render_from_plan(plan_payload, output_path=output_path)
    else:
        request_payload = load_json_payload(args.request)
        request_obj = normalize_request_payload(request_payload)
        result = renderer.render_from_request(request_obj, output_path=output_path)

    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))
    print(f"Generated PPTX: {result.output_path}")


if __name__ == "__main__":
    main()
