import argparse
import json

from slides_service import SlidesService


def main() -> None:
    parser = argparse.ArgumentParser(description="Run slides -> PPTX generation pipeline.")
    parser.add_argument("--project-id", required=True, help="Project id to process.")
    parser.add_argument(
        "--mode",
        choices=["json", "topic"],
        default="json",
        help="`json`: use existing json_files. `topic`: rebuild json from topic content.",
    )
    parser.add_argument(
        "--topic",
        action="append",
        dest="topics",
        default=None,
        help="Optional topic title filter. Can be provided multiple times.",
    )
    args = parser.parse_args()

    generation_client = object() if args.mode == "topic" else None
    service = SlidesService(generation_client=generation_client)
    result = service.generate_slides(project_id=str(args.project_id), topic_titles=args.topics)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

