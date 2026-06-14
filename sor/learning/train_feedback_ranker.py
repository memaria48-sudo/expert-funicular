"""Train the SCIP neural feedback ranker.

This is an offline training step. It reads explicit dashboard feedback,
encodes each feedback item as a fixed-width candidate vector, trains a small
MLP with BCE loss, and stores a bounded reranker model for the optimizer.
"""

from __future__ import annotations

import argparse
import json
import random
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence

import numpy as np

try:
    import torch
    import torch.nn as nn
    from torch.utils.data import DataLoader, TensorDataset
except Exception as exc:  # pragma: no cover
    torch = None
    nn = None
    DataLoader = None
    TensorDataset = None
    TORCH_IMPORT_ERROR = exc
else:
    TORCH_IMPORT_ERROR = None

from .build_training_data import build_training_dataset
from .feature_encoder import CandidateFeatureEncoder
from .model_registry import metrics_path as default_metrics_path
from .model_registry import model_path as default_model_path
from .model_registry import project_root_from_here
from .neural_reranker import FeedbackRanker


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _label_counts(examples: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for example in examples:
        label = str(example.get("label") or "")
        counts[label] = counts.get(label, 0) + 1
    return counts


def _load_training_dataset(path: Path) -> list[dict[str, Any]]:
    examples: list[dict[str, Any]] = []
    if not path.exists():
        return examples
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            raw = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(raw, dict) and "target" in raw:
            examples.append(raw)
    return examples


def train_feedback_ranker(
    *,
    project_root: Path | None = None,
    model_output_path: Path | None = None,
    metrics_output_path: Path | None = None,
    min_examples: int = 30,
    epochs: int = 80,
    batch_size: int = 16,
    learning_rate: float = 0.001,
    seed: int = 42,
) -> dict[str, Any]:
    root = project_root or project_root_from_here()
    model_out = model_output_path or default_model_path(root)
    metrics_out = metrics_output_path or default_metrics_path(root)
    dataset_result = build_training_dataset(root)
    dataset_path = Path(str(dataset_result.get("path") or ""))
    examples = _load_training_dataset(dataset_path)

    metrics_base: dict[str, Any] = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "dataset": dataset_result,
        "n_examples": len(examples),
        "label_counts": _label_counts(examples),
        "model_path": str(model_out),
        "status": "",
    }

    if torch is None or nn is None or DataLoader is None or TensorDataset is None:
        metrics = {
            **metrics_base,
            "status": "skipped",
            "reason": "torch_unavailable",
            "error": str(TORCH_IMPORT_ERROR),
        }
        _write_json(metrics_out, metrics)
        return metrics

    if len(examples) < min_examples:
        metrics = {
            **metrics_base,
            "status": "skipped",
            "reason": "not_enough_feedback",
            "min_examples": min_examples,
        }
        _write_json(metrics_out, metrics)
        return metrics

    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)

    encoder = CandidateFeatureEncoder()
    x = np.stack([encoder.encode_candidate(example) for example in examples]).astype(np.float32)
    y = np.array([float(example.get("target", 0.0)) for example in examples], dtype=np.float32).reshape(-1, 1)

    indices = np.arange(len(examples))
    np.random.shuffle(indices)
    split = max(1, int(len(indices) * 0.82))
    train_idx = indices[:split]
    val_idx = indices[split:] if len(indices[split:]) else indices[: min(8, len(indices))]

    train_ds = TensorDataset(torch.tensor(x[train_idx]), torch.tensor(y[train_idx]))
    loader = DataLoader(train_ds, batch_size=max(1, int(batch_size)), shuffle=True)

    model = FeedbackRanker(input_dim=x.shape[1])
    optimizer = torch.optim.Adam(model.parameters(), lr=float(learning_rate), weight_decay=1e-4)
    loss_fn = nn.BCELoss()
    losses: list[float] = []

    model.train()
    for _ in range(max(1, int(epochs))):
        epoch_losses: list[float] = []
        for xb, yb in loader:
            optimizer.zero_grad()
            pred = model(xb)
            loss = loss_fn(pred, yb)
            loss.backward()
            optimizer.step()
            epoch_losses.append(float(loss.item()))
        losses.append(float(np.mean(epoch_losses)) if epoch_losses else 0.0)

    model.eval()
    with torch.no_grad():
        val_pred = model(torch.tensor(x[val_idx])).numpy().reshape(-1)
    val_target = y[val_idx].reshape(-1)
    val_mae = float(np.mean(np.abs(val_pred - val_target))) if len(val_target) else 0.0
    val_brier = float(np.mean((val_pred - val_target) ** 2)) if len(val_target) else 0.0

    model_out.parent.mkdir(parents=True, exist_ok=True)
    torch.save(
        {
            "model_state_dict": model.state_dict(),
            "input_dim": int(x.shape[1]),
            "n_examples": len(examples),
            "encoder": encoder.to_config(),
            "label_counts": _label_counts(examples),
            "trained_at": datetime.now(timezone.utc).isoformat(),
        },
        model_out,
    )

    metrics = {
        **metrics_base,
        "status": "trained",
        "input_dim": int(x.shape[1]),
        "epochs": int(epochs),
        "batch_size": int(batch_size),
        "learning_rate": float(learning_rate),
        "train_examples": int(len(train_idx)),
        "validation_examples": int(len(val_idx)),
        "loss_first": round(losses[0], 6) if losses else None,
        "final_loss": round(losses[-1], 6) if losses else None,
        "loss_history_tail": [round(v, 6) for v in losses[-10:]],
        "validation_mae": round(val_mae, 6),
        "validation_brier": round(val_brier, 6),
    }
    _write_json(metrics_out, metrics)
    return metrics


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Train the SCIP neural feedback ranker.")
    parser.add_argument("--project-root", type=Path, default=project_root_from_here())
    parser.add_argument("--min-examples", type=int, default=30)
    parser.add_argument("--epochs", type=int, default=80)
    parser.add_argument("--batch-size", type=int, default=16)
    parser.add_argument("--learning-rate", type=float, default=0.001)
    args = parser.parse_args(argv)
    result = train_feedback_ranker(
        project_root=args.project_root,
        min_examples=args.min_examples,
        epochs=args.epochs,
        batch_size=args.batch_size,
        learning_rate=args.learning_rate,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
