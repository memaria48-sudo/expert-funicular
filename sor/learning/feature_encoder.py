"""Feature encoding for the SCIP neural feedback ranker.

The encoder is intentionally fixed-width and dependency-light. It combines a
small numeric spine with hashed categorical features so new actors, topics and
subroles can appear without retraining an explicit vocabulary.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from typing import Any, Iterable

import numpy as np


NUMERIC_FIELDS = (
    "base_score",
    "score",
    "raw_score",
    "rule_final_score",
    "static_support_score",
    "live_signal_boost",
    "evidence_count",
    "live_event_count",
    "company_evidence_count",
    "academic_evidence_count",
    "freshness_days",
    "critical_action_penalty",
    "memory_adjustment",
    "explicit_feedback_delta",
    "feedback_count",
    "word_memory_hits",
    "profit_anchor_count",
    "implementation_anchor_count",
    "academic_support_count",
    "context_actor_count",
)

SCORE_FIELDS = {
    "base_score",
    "score",
    "raw_score",
    "rule_final_score",
    "static_support_score",
    "critical_action_penalty",
    "memory_adjustment",
    "explicit_feedback_delta",
}

COUNT_FIELDS = {
    "evidence_count",
    "live_event_count",
    "company_evidence_count",
    "academic_evidence_count",
    "feedback_count",
    "word_memory_hits",
}


def normalize_text(value: Any) -> str:
    text = str(value or "").replace("ß", "ss").replace("ẞ", "ss")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return " ".join(re.sub(r"[^a-zA-Z0-9]+", " ", text.lower()).split())


def stable_hash(value: str, buckets: int) -> int:
    digest = hashlib.sha256(normalize_text(value).encode("utf-8")).hexdigest()
    return int(digest, 16) % buckets


def as_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def list_from_any(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, tuple):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str) and value.strip():
        parts = re.split(r"\s+(?:x|\|)\s+|[,;]", value)
        return [part.strip() for part in parts if part.strip()]
    return []


def sorted_join(values: Iterable[Any]) -> str:
    return " x ".join(sorted(normalize_text(value) for value in values if str(value).strip()))


class CandidateFeatureEncoder:
    def __init__(self, hash_buckets: int = 96):
        self.hash_buckets = int(hash_buckets)

    @property
    def input_dim(self) -> int:
        return len(NUMERIC_FIELDS) + 8 + self.hash_buckets * 7

    def _hash_feature(self, value: str) -> np.ndarray:
        vec = np.zeros(self.hash_buckets, dtype=np.float32)
        if value:
            vec[stable_hash(value, self.hash_buckets)] = 1.0
        return vec

    def _numeric_value(self, candidate: dict[str, Any], field: str) -> float:
        features = candidate.get("features") if isinstance(candidate.get("features"), dict) else {}
        value = candidate.get(field, features.get(field))
        number = as_float(value)
        if field in SCORE_FIELDS:
            return max(-1.0, min(1.5, number / 100.0))
        if field in COUNT_FIELDS:
            return max(0.0, min(2.0, number / 20.0))
        if field == "freshness_days":
            return max(0.0, min(2.0, number / 365.0))
        return max(-2.0, min(2.0, number))

    def encode_candidate(self, candidate: dict[str, Any]) -> np.ndarray:
        actors = list_from_any(candidate.get("actors") or candidate.get("members") or candidate.get("pair"))
        actor_types = candidate.get("actor_types") if isinstance(candidate.get("actor_types"), dict) else {}
        actor_roles = list_from_any(candidate.get("actor_roles"))
        if not actor_roles and actor_types:
            actor_roles = [str(actor_types.get(actor, "")).strip() for actor in actors if str(actor_types.get(actor, "")).strip()]
        actor_subroles = list_from_any(candidate.get("actor_subroles"))
        topic = str(candidate.get("topic") or candidate.get("convening_theme") or "")
        cluster = str(candidate.get("cluster") or candidate.get("cluster_label") or candidate.get("concrete_clusters") or "")
        opportunity = str(candidate.get("opportunity_id") or candidate.get("opportunity_key") or "")
        reason_category = str(candidate.get("reason_category") or candidate.get("feedback_reason_category") or "")
        feedback_target = str(candidate.get("feedback_target") or candidate.get("feedback_dimension") or "")
        words = list_from_any(candidate.get("words"))
        reason = str(candidate.get("reason") or candidate.get("editorial_justification") or "")

        numeric = np.array([self._numeric_value(candidate, field) for field in NUMERIC_FIELDS], dtype=np.float32)
        flags = np.array([
            1.0 if candidate.get("selected") in {True, 1, "1", "true", "True"} else 0.0,
            1.0 if candidate.get("academic_flag") or as_float(candidate.get("academic_support_count")) > 0 else 0.0,
            1.0 if candidate.get("academic_only_flag") else 0.0,
            1.0 if candidate.get("company_need_confirmed") else 0.0,
            1.0 if as_float(candidate.get("live_event_count")) > 0 else 0.0,
            1.0 if as_float(candidate.get("critical_action_penalty")) <= 0 else 0.0,
            1.0 if len(actors) >= 2 else 0.0,
            1.0 if actor_roles and len(set(actor_roles)) >= 2 else 0.0,
        ], dtype=np.float32)

        hashed = np.concatenate([
            self._hash_feature(sorted_join(actors)),
            self._hash_feature(sorted_join(actor_roles)),
            self._hash_feature(sorted_join(actor_subroles)),
            self._hash_feature(topic),
            self._hash_feature(cluster),
            self._hash_feature(opportunity),
            self._hash_feature(" ".join([reason_category, feedback_target, reason, " ".join(words[:20])])),
        ])
        return np.concatenate([numeric, flags, hashed]).astype(np.float32)

    def to_config(self) -> dict[str, Any]:
        return {
            "hash_buckets": self.hash_buckets,
            "input_dim": self.input_dim,
            "numeric_fields": list(NUMERIC_FIELDS),
        }
