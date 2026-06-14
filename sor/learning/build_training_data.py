"""Build a training dataset from SCIP dashboard feedback."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Sequence

from .model_registry import project_root_from_here, training_dataset_path


LABEL_TO_TARGET = {
    "useful": 1.0,
    "interesting_but_weak": 0.45,
    "good_topic_wrong_actors": 0.30,
    "wrong_actors": 0.30,
    "wrong_connection": 0.10,
    "not_relevant": 0.0,
}


def normalize_label(label: Any) -> str:
    return str(label or "").strip().lower().replace("-", "_").replace(" ", "_")


def feedback_file(project_root: Path) -> Path:
    return project_root / "data" / "feedback" / "scip_feedback.jsonl"


def feedback_to_example(record: dict[str, Any]) -> dict[str, Any] | None:
    label = normalize_label(record.get("label") or record.get("human_feedback_label"))
    if label not in LABEL_TO_TARGET:
        return None
    candidate = dict(record)
    features = candidate.get("features")
    if isinstance(features, dict):
        for key, value in features.items():
            candidate.setdefault(key, value)
    candidate["label"] = label
    candidate["target"] = LABEL_TO_TARGET[label]
    return candidate


def load_feedback_examples(project_root: Path) -> list[dict[str, Any]]:
    path = feedback_file(project_root)
    if not path.exists():
        return []
    examples: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            raw = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(raw, dict):
            example = feedback_to_example(raw)
            if example is not None:
                examples.append(example)
    return examples


def build_training_dataset(project_root: Path | None = None, output_path: Path | None = None) -> dict[str, Any]:
    root = project_root or project_root_from_here()
    out = output_path or training_dataset_path(root)
    examples = load_feedback_examples(root)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as handle:
        for example in examples:
            handle.write(json.dumps(example, ensure_ascii=False) + "\n")
    labels: dict[str, int] = {}
    for example in examples:
        labels[example["label"]] = labels.get(example["label"], 0) + 1
    return {
        "status": "written",
        "path": str(out),
        "n_examples": len(examples),
        "labels": labels,
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build SCIP feedback training data.")
    parser.add_argument("--project-root", type=Path, default=project_root_from_here())
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args(argv)
    result = build_training_dataset(args.project_root, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
