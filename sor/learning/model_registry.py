"""Filesystem paths for SOR learning artifacts."""

from __future__ import annotations

from pathlib import Path


def project_root_from_here() -> Path:
    return Path(__file__).resolve().parents[2]


def learning_dir(project_root: Path | None = None) -> Path:
    root = project_root or project_root_from_here()
    return root / "data" / "learning"


def models_dir(project_root: Path | None = None) -> Path:
    root = project_root or project_root_from_here()
    return root / "data" / "models"


def training_dataset_path(project_root: Path | None = None) -> Path:
    return learning_dir(project_root) / "scip_training_dataset.jsonl"


def model_path(project_root: Path | None = None) -> Path:
    return models_dir(project_root) / "scip_feedback_ranker.pt"


def metrics_path(project_root: Path | None = None) -> Path:
    return models_dir(project_root) / "scip_feedback_ranker_metrics.json"
