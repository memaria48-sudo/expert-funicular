"""Build a training dataset from SCIP dashboard feedback."""

from __future__ import annotations

import argparse
import json
import sys
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


def _target_from_adjustment(adjustment: float, feedback_count: int) -> float:
    """Turn the bounded rule/human feedback delta into a NN target.

    Neutral cards stay at 0.50 so the model sees the whole current candidate
    space without inventing negative labels for unseen opportunities.
    """
    if feedback_count <= 0:
        return 0.50
    return max(0.05, min(0.95, 0.50 + (float(adjustment) / 48.0)))


def load_current_pool_examples(project_root: Path) -> list[dict[str, Any]]:
    """Build one training example for each current SCIP optimization-pool card."""
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    import pandas as pd
    import zirp_meeting_optimizer as optimizer

    workbook = project_root / "zirp(3).xlsx"
    patterns = optimizer.load_static_candidate_patterns_from_excel(workbook, pd.DataFrame())
    examples: list[dict[str, Any]] = []
    for pattern in patterns:
        base_score = optimizer.rule_pattern_score(pattern.score, pattern)
        payload = optimizer.neural_candidate_payload(base_score, pattern)
        human_delta, human_meta = optimizer.explicit_human_feedback_adjustment_for_pattern(pattern)
        memory_delta, memory_meta = optimizer.feedback_adjustment_for_pattern(base_score, pattern)
        exact_count = int(human_meta.get("human_feedback_exact_count", 0) or 0)
        similar_count = int(human_meta.get("human_feedback_similar_count", 0) or 0)
        human_count = int(human_meta.get("human_feedback_count", 0) or 0)
        memory_count = int(memory_meta.get("feedback_evidence_count", 0) or 0)
        total_delta = float(human_delta or 0.0) + float(memory_delta or 0.0)
        if exact_count:
            label = "pool_exact_feedback"
        elif similar_count:
            label = "pool_similar_feedback"
        elif memory_count:
            label = "pool_word_memory"
        else:
            label = "pool_neutral"
        payload.update({
            "label": label,
            "target": _target_from_adjustment(total_delta, human_count + memory_count),
            "pool_training_example": True,
            "pool_pattern_id": pattern.pattern_id,
            "human_feedback_adjustment": human_delta,
            "memory_adjustment": memory_delta,
            "human_feedback_count": human_count,
            "human_feedback_exact_count": exact_count,
            "human_feedback_similar_count": similar_count,
            "feedback_evidence_count": memory_count,
        })
        features = dict(payload.get("features") or {})
        features.update({
            "memory_adjustment": float(memory_delta or 0.0),
            "explicit_feedback_delta": float(human_delta or 0.0),
            "feedback_count": float(human_count),
            "word_memory_hits": float(memory_meta.get("word_memory_hits", 0) or 0),
        })
        payload["features"] = features
        examples.append(payload)
    return examples


def build_training_dataset(project_root: Path | None = None, output_path: Path | None = None) -> dict[str, Any]:
    root = project_root or project_root_from_here()
    out = output_path or training_dataset_path(root)
    explicit_examples = load_feedback_examples(root)
    pool_examples = load_current_pool_examples(root)
    examples = explicit_examples + pool_examples
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
        "explicit_feedback_examples": len(explicit_examples),
        "current_pool_examples": len(pool_examples),
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
