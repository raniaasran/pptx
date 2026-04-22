from __future__ import annotations

import json

from mini_layout_engine.engine.planning_engine import PlanningEngine


def build_example_request() -> dict:
    return {
        "request_id": "demo_request_001",
        "design_family_id": "modern_v1",
        "slides": [
            {
                "layout_id": "story_points_image",
                "content": {
                    "title": "AI Triage Snapshot",
                    "intro_paragraphs": "A support desk introduced AI triage to classify incoming tickets.",
                    "points": [
                        "Routing rules cut first-response latency by 31%.",
                        "Escalation quality improved for urgent incidents.",
                        "Weekly reviews kept confidence high across teams.",
                    ],
                    "image": "ai_image",
                },
            },
            {
                "layout_id": "story_points_image",
                "content": {
                    "title": "ML Quality Loop",
                    "intro_paragraphs": [
                        "The team established a repeatable model quality process.",
                        "Monitoring signals were reviewed every sprint.",
                    ],
                    "points": "Baseline drift alerts reduced silent failures.",
                    "closing_paragraph": "Stakeholders received monthly quality summaries.",
                    "image": "ml_image",
                },
            },
            {
                "layout_id": "dual_cards",
                "content": {
                    "title": "Rule-based vs Model-based Routing",
                    "left_title": "Rule-based",
                    "right_title": "Model-based",
                    "left_points": [
                        "Fast and predictable for narrow cases.",
                        "Needs manual updates when patterns change.",
                    ],
                    "right_points": [
                        "Scales better with language variability.",
                        "Requires monitoring and model governance.",
                    ],
                    "footer_note": "Choose by risk profile and operational maturity.",
                },
            },
            {
                "layout_id": "unknown_layout",
                "content": {
                    "title": "Intentional Validation Example"
                },
            },
        ],
    }


def main() -> None:
    engine = PlanningEngine.from_default_specs()
    result = engine.plan_deck(build_example_request())
    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
