"""Neural feedback reranker for SCIP candidates.

The model is deliberately small and bounded. It predicts p(useful) from the
same structured candidate payload that the dashboard feedback writes, then
turns that probability into a limited score delta. If Torch or a trained model
is unavailable, the layer returns a neutral delta and the optimizer behaves
exactly like the rule/feedback-memory system.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

try:
    import torch
    import torch.nn as nn
except Exception:  # pragma: no cover - production fallback when torch is absent
    torch = None
    nn = None

from .feature_encoder import CandidateFeatureEncoder
from .model_registry import model_path as default_model_path


class FeedbackRanker(nn.Module if nn is not None else object):
    def __init__(self, input_dim: int):
        if nn is None:
            raise RuntimeError("PyTorch is not available")
        super().__init__()
        self.net = nn.Sequential(
            nn.Linear(input_dim, 64),
            nn.ReLU(),
            nn.Dropout(0.20),
            nn.Linear(64, 32),
            nn.ReLU(),
            nn.Linear(32, 1),
            nn.Sigmoid(),
        )

    def forward(self, x):  # type: ignore[no-untyped-def]
        return self.net(x)


class NeuralFeedbackReranker:
    def __init__(
        self,
        model_path: str | Path,
        *,
        enabled: bool = True,
        default_probability: float = 0.5,
        max_delta: float = 8.0,
    ):
        self.model_path = Path(model_path)
        self.enabled = bool(enabled)
        self.default_probability = float(default_probability)
        self.max_delta = float(max_delta)
        self.encoder = CandidateFeatureEncoder()
        self.model: FeedbackRanker | None = None
        self.input_dim: int = self.encoder.input_dim
        self.n_examples: int = 0
        self.status: str = "disabled" if not self.enabled else "not_loaded"
        self._cache: dict[str, dict[str, Any]] = {}
        self._load_model()

    @classmethod
    def from_project_root(
        cls,
        project_root: Path,
        *,
        enabled: bool = True,
        max_delta: float = 8.0,
    ) -> "NeuralFeedbackReranker":
        return cls(default_model_path(project_root), enabled=enabled, max_delta=max_delta)

    def _load_model(self) -> None:
        if not self.enabled:
            return
        if torch is None or nn is None:
            self.status = "torch_unavailable"
            return
        if not self.model_path.exists():
            self.status = "model_missing"
            return
        try:
            checkpoint = torch.load(self.model_path, map_location="cpu")
            self.input_dim = int(checkpoint.get("input_dim") or self.encoder.input_dim)
            encoder_config = checkpoint.get("encoder") or {}
            if isinstance(encoder_config, dict) and encoder_config.get("hash_buckets"):
                self.encoder = CandidateFeatureEncoder(hash_buckets=int(encoder_config["hash_buckets"]))
            self.model = FeedbackRanker(self.input_dim)
            self.model.load_state_dict(checkpoint["model_state_dict"])
            self.model.eval()
            self.n_examples = int(checkpoint.get("n_examples") or 0)
            self.status = "loaded"
        except Exception as exc:
            self.model = None
            self.status = f"load_failed:{type(exc).__name__}"

    def dynamic_max_delta(self) -> float:
        if self.n_examples < 100:
            return min(self.max_delta, 3.0)
        if self.n_examples < 300:
            return min(self.max_delta, 5.0)
        if self.n_examples < 700:
            return min(self.max_delta, 8.0)
        return self.max_delta

    def _cache_key(self, candidate: dict[str, Any]) -> str:
        identity = "|".join(
            str(candidate.get(key, ""))
            for key in ("recommendation_id", "opportunity_id", "pair", "topic", "score", "base_score")
        )
        return hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16]

    def predict_probability(self, candidate: dict[str, Any]) -> float:
        if self.model is None or torch is None:
            return self.default_probability
        x = self.encoder.encode_candidate(candidate)
        if x.shape[0] != self.input_dim:
            return self.default_probability
        with torch.no_grad():
            xt = torch.tensor(x, dtype=torch.float32).unsqueeze(0)
            return float(self.model(xt).item())

    def score_delta(self, candidate: dict[str, Any]) -> dict[str, Any]:
        key = self._cache_key(candidate)
        if key in self._cache:
            return dict(self._cache[key])
        probability = self.predict_probability(candidate)
        max_delta = self.dynamic_max_delta()
        raw_delta = (probability - 0.5) * 2.0 * max_delta
        delta = max(-max_delta, min(max_delta, raw_delta))
        result = {
            "nn_probability_useful": round(probability, 4),
            "nn_delta": round(delta, 3),
            "nn_model_examples": self.n_examples,
            "nn_max_delta": round(max_delta, 3),
            "nn_model_status": self.status,
        }
        self._cache[key] = dict(result)
        return result
