from __future__ import annotations

import csv
import hashlib
import html
import importlib.util
import json
import math
import os
import re
import subprocess
import unicodedata
import urllib.error
import urllib.request
from collections import Counter, defaultdict
from enum import Enum
from dataclasses import dataclass
from datetime import datetime, timedelta
from itertools import combinations
from pathlib import Path
from typing import Any, Optional

import pandas as pd

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

try:
    from sor.app.nn.scip_feedback import SCIPFeedbackLearner, normalize_words, stable_recommendation_id
except Exception:
    SCIPFeedbackLearner = None
    normalize_words = None
    stable_recommendation_id = None

try:
    from sor.learning.neural_reranker import NeuralFeedbackReranker
except Exception:
    NeuralFeedbackReranker = None

try:
    from sor.app.network_config import load_network_config
except Exception:
    load_network_config = None


PROJECT_DIR = Path(__file__).resolve().parent
REPORT_DIR = PROJECT_DIR / "zirp_berichte"
SCIP_ROOT = Path(r"C:\Users\memar\SCIPOptSuite 10.0.0\app stuff")
SCIP_EXE = SCIP_ROOT / "bin" / "scip.exe"
NETWORK_NAME = os.getenv("SOR_NETWORK", os.getenv("ZIRP_NETWORK", "zirp")).strip() or "zirp"

LOOKBACK_DAYS = 14
RECENT_PRIORITY_DAYS = 7
RECENT_SIGNAL_WEIGHT = 1.35
PAIR_ONLY = True
MIN_CLUSTER_SIZE = 2
MAX_CLUSTER_SIZE = 2
MAX_MEETINGS = 10
MAX_PUBLIC_CARDS = 10
MIN_PUBLIC_CARDS = 10
MIN_CANDIDATE_PATTERNS = 10
MAX_MEETINGS_PER_MEMBER = 1
MAX_CANDIDATE_PATTERNS = 200
MIN_FINAL_PAIR_SCORE = 35
REQUIRE_IMPLEMENTATION_ANCHOR = True
MAX_SELECTED_PER_CLUSTER = 2
MAX_SELECTED_PER_OPPORTUNITY = 1
USE_OLLAMA_FOR_CONVENING_TEXT = os.getenv("ZIRP_CONVENING_AI", "1").strip().lower() not in {"0", "false", "no", "off"}
OLLAMA_CONVENING_TIMEOUT = 60
OLLAMA_CONVENING_MAX_TOKENS = 360
MAX_MODEL_CONVENING_CARDS = 6
USE_STATIC_CANDIDATES = os.getenv("ZIRP_USE_STATIC_CANDIDATES", "1").strip().lower() not in {"0", "false", "no", "off"}
# v2: prefer the latest static/crawler workbook if it exists, otherwise fall back.
_DEFAULT_STATIC_CANDIDATE_PATH = PROJECT_DIR / "zirp(3).xlsx"
if not _DEFAULT_STATIC_CANDIDATE_PATH.exists():
    _DEFAULT_STATIC_CANDIDATE_PATH = PROJECT_DIR / "sor_scip_full_static_zirp_dataset_updated.xlsx"
if not _DEFAULT_STATIC_CANDIDATE_PATH.exists():
    _DEFAULT_STATIC_CANDIDATE_PATH = PROJECT_DIR / "sor_scip_full_static_zirp_dataset.xlsx"
STATIC_CANDIDATE_PATH = Path(os.getenv("ZIRP_STATIC_CANDIDATE_PATH", str(_DEFAULT_STATIC_CANDIDATE_PATH)))
STATIC_CANDIDATE_BASE_PATH = Path(os.getenv("ZIRP_STATIC_CANDIDATE_BASE_PATH", str(PROJECT_DIR / "sor_scip_full_static_zirp_dataset.xlsx")))
STATIC_CANDIDATE_SHEET = os.getenv("ZIRP_STATIC_CANDIDATE_SHEET", "Candidate_Patterns")
STATIC_ACTOR_VECTOR_SHEET = os.getenv("ZIRP_STATIC_ACTOR_VECTOR_SHEET", "Actor_Vector")
STATIC_PAIR_VECTOR_SHEET = os.getenv("ZIRP_STATIC_PAIR_VECTOR_SHEET", "Pair_Vector")
STATIC_PAIR_TOPIC_VECTOR_SHEET = os.getenv("ZIRP_STATIC_PAIR_TOPIC_VECTOR_SHEET", "Pair_Topic_Vector")
STATIC_PAIR_EVIDENCE_VECTOR_SHEET = os.getenv("ZIRP_STATIC_PAIR_EVIDENCE_VECTOR_SHEET", "Pair_Evidence_Vector")
STATIC_CARD_SHEET = os.getenv("ZIRP_STATIC_CARD_SHEET", "Opportunity_Cards")
STATIC_FRONTIER_SHEET = os.getenv("ZIRP_STATIC_FRONTIER_SHEET", "SCIP_Frontier")
STATIC_SCIP_INPUT_SHEET = os.getenv("ZIRP_STATIC_SCIP_INPUT_SHEET", "SCIP_Optimization_Pool")
REFRESH_STATIC_SCIP_POOL = os.getenv("ZIRP_REFRESH_SCIP_POOL", "1").strip().lower() not in {"0", "false", "no", "off"}
SCIP_POOL_HIGH_SCORE_COUNT = int(os.getenv("ZIRP_SCIP_POOL_HIGH_SCORE_COUNT", "650") or "650")
SCIP_POOL_FEEDBACK_NEIGHBOR_COUNT = int(os.getenv("ZIRP_SCIP_POOL_FEEDBACK_NEIGHBOR_COUNT", "100") or "100")
SCIP_POOL_EXPLORATION_COUNT = int(os.getenv("ZIRP_SCIP_POOL_EXPLORATION_COUNT", "50") or "50")
SCIP_POOL_TARGET_COUNT = int(os.getenv("ZIRP_SCIP_POOL_TARGET_COUNT", "800") or "800")
_DEFAULT_STATIC_OPTIMIZER_PATH = STATIC_CANDIDATE_PATH
if not _DEFAULT_STATIC_OPTIMIZER_PATH.exists():
    _DEFAULT_STATIC_OPTIMIZER_PATH = PROJECT_DIR / "zirp_nested_vector_space.xlsx"
STATIC_OPTIMIZER_PATH = Path(os.getenv("ZIRP_STATIC_OPTIMIZER_PATH", str(_DEFAULT_STATIC_OPTIMIZER_PATH)))
STATIC_FULL_PAIRING_SHEET = os.getenv("ZIRP_STATIC_FULL_PAIRING_SHEET", "All_Pairings_5253")
USE_STATIC_FULL_PAIRING_UNIVERSE = os.getenv("ZIRP_USE_FULL_PAIRING_UNIVERSE", "1").strip().lower() not in {"0", "false", "no", "off"}
STATIC_CLUSTER_SHEET = os.getenv("ZIRP_STATIC_CLUSTER_SHEET", "Static_Clusters")
STATIC_MEMBER_SHEET = os.getenv("ZIRP_STATIC_MEMBER_SHEET", "All_ZIRP_Members")
STATIC_MEMBER_CLUSTER_MATRIX_SHEET = os.getenv("ZIRP_STATIC_MEMBER_CLUSTER_MATRIX_SHEET", "Member_Cluster_Matrix")
STATIC_MEMBER_CONSTRAINT_SHEET = os.getenv("ZIRP_STATIC_MEMBER_CONSTRAINT_SHEET", "Member_Constraints")
STATIC_CLUSTER_CONSTRAINT_SHEET = os.getenv("ZIRP_STATIC_CLUSTER_CONSTRAINT_SHEET", "Cluster_Constraints")
STATIC_CANDIDATE_LIMIT = int(os.getenv("ZIRP_STATIC_CANDIDATE_LIMIT", "0") or "0")
STATIC_PATCH_MIN_CANDIDATES = int(os.getenv("ZIRP_STATIC_PATCH_MIN_CANDIDATES", "50") or "50")
STATIC_REMOVED_MEMBER_NAMES_RAW = [
    name for name in os.getenv("ZIRP_STATIC_REMOVED_MEMBERS", "").split("|") if name.strip()
]
STATIC_REMOVED_MEMBERS: set[str] = set()
STATIC_MEMBER_LIMITS: dict[str, int] = {}
STATIC_CLUSTER_LIMITS: dict[str, int] = {}
STATIC_ROLE_MASS_LIMITS: dict[str, float] = {}
STATIC_ROLE_INVERSE_WEIGHTS: dict[str, float] = {}
STATIC_MEMBER_ROLE_WEIGHTS: dict[str, tuple[str, float]] = {}
STATIC_SUBROLE_MASS_LIMITS: dict[str, float] = {}
STATIC_SUBROLE_INVERSE_WEIGHTS: dict[str, float] = {}
STATIC_MEMBER_SUBROLE_WEIGHTS: dict[str, tuple[str, float]] = {}
STATIC_MAX_ACADEMIC_MATCHES: int = int(os.getenv("ZIRP_MAX_ACADEMIC_MATCHES", "3") or "3")
STATIC_MAX_ACADEMIC_ACTOR_SLOTS: int = int(os.getenv("ZIRP_MAX_ACADEMIC_ACTOR_SLOTS", "4") or "4")
STATIC_MIN_NON_ACADEMIC_PAIRS: int = int(os.getenv("ZIRP_MIN_NON_ACADEMIC_PAIRS", "7") or "7")
STATIC_MAX_WEAK_EVIDENCE_MATCHES: int = int(os.getenv("ZIRP_MAX_WEAK_EVIDENCE_MATCHES", "3") or "3")
STATIC_MAX_LOW_SCORE_MATCHES: int = int(os.getenv("ZIRP_MAX_LOW_SCORE_MATCHES", "2") or "2")
STATIC_MIN_GEO_PROXIMITY_SUM: float = float(os.getenv("ZIRP_MIN_GEO_PROXIMITY_SUM", "5.5") or "5.5")
STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS: int = int(os.getenv("ZIRP_MAX_ACADEMIC_ACADEMIC_PAIRS", "0") or "0")

FEEDBACK_LEARNER = SCIPFeedbackLearner.from_project_root(PROJECT_DIR) if SCIPFeedbackLearner is not None else None
USE_NEURAL_FEEDBACK_RANKER = os.getenv("SOR_NEURAL_FEEDBACK_RANKER", "1").strip().lower() not in {"0", "false", "no", "off"}
NEURAL_FEEDBACK_RERANKER = (
    NeuralFeedbackReranker.from_project_root(PROJECT_DIR, enabled=USE_NEURAL_FEEDBACK_RANKER)
    if NeuralFeedbackReranker is not None
    else None
)
OPPORTUNITY_HISTORY_PATH = PROJECT_DIR / "data" / "feedback" / "opportunity_history.json"
DYNAMIC_TERM_MEMORY_PATH = PROJECT_DIR / "data" / "feedback" / "dynamic_term_memory.json"


def normalize_search_text(value: Any) -> str:
    text = str(value or "").lower()
    for bad, good in {
        "\u00c3\u00a4": "\u00e4",
        "\u00c3\u00b6": "\u00f6",
        "\u00c3\u00bc": "\u00fc",
        "\u00c3\u009f": "\u00df",
        "\u00e2\u20ac\u201c": "-",
        "\u00e2\u20ac\u201d": "-",
        "\u00c2\u00b7": " ",
    }.items():
        text = text.replace(bad, good)
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("\u00df", "ss")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def network_member_key(member_name: str) -> str:
    return f" {normalize_search_text(member_name)} "


# Normalize after normalize_search_text is available.
STATIC_REMOVED_MEMBERS = {
    normalize_search_text(name)
    for name in STATIC_REMOVED_MEMBER_NAMES_RAW
    if normalize_search_text(name)
}


def _load_network_settings() -> tuple[dict[str, Any], str, dict[str, str]]:
    if load_network_config is None:
        return {}, "", {}
    try:
        network_config = load_network_config(NETWORK_NAME, PROJECT_DIR)
    except Exception as exc:
        print(f"Network config: using built-in defaults ({exc})")
        return {}, "", {}
    print(f"Network config: loaded {network_config.name} from {network_config.network_dir}")
    member_roles: dict[str, str] = {}
    valid_roles = {"profit_anchor", "implementation_anchor", "academic_support", "context_actor"}
    if not network_config.members.empty and {"name", "actor_type"}.issubset(network_config.members.columns):
        for _, row in network_config.members.iterrows():
            name = str(row.get("name", "") or "").strip()
            role = str(row.get("actor_type", "") or "").strip()
            if name and role in valid_roles:
                member_roles[network_member_key(name)] = role
    return dict(network_config.optimizer or {}), str(network_config.network_dir), member_roles


NETWORK_OPTIMIZER, NETWORK_CONFIG_DIR, NETWORK_MEMBER_ROLES = _load_network_settings()


def optimizer_bool(name: str, default: bool) -> bool:
    value = NETWORK_OPTIMIZER.get(name, default)
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() not in {"0", "false", "no", "off"}


def optimizer_int(name: str, default: int) -> int:
    try:
        return int(NETWORK_OPTIMIZER.get(name, default))
    except Exception:
        return default


PAIR_ONLY = optimizer_bool("pair_only", PAIR_ONLY)
MIN_CLUSTER_SIZE = optimizer_int("min_cluster_size", MIN_CLUSTER_SIZE)
MAX_CLUSTER_SIZE = optimizer_int("max_cluster_size", MAX_CLUSTER_SIZE)
MAX_MEETINGS = optimizer_int("max_matches", MAX_MEETINGS)
MAX_PUBLIC_CARDS = optimizer_int("max_public_cards", MAX_PUBLIC_CARDS)
MIN_PUBLIC_CARDS = optimizer_int("min_public_cards", MIN_PUBLIC_CARDS)
MIN_CANDIDATE_PATTERNS = optimizer_int("min_candidate_patterns", MIN_CANDIDATE_PATTERNS)
MAX_MEETINGS_PER_MEMBER = optimizer_int("max_matches_per_member", MAX_MEETINGS_PER_MEMBER)
MAX_CANDIDATE_PATTERNS = optimizer_int("max_candidate_patterns", MAX_CANDIDATE_PATTERNS)
MIN_FINAL_PAIR_SCORE = optimizer_int("min_final_pair_score", MIN_FINAL_PAIR_SCORE)
REQUIRE_IMPLEMENTATION_ANCHOR = optimizer_bool("require_implementation_anchor", REQUIRE_IMPLEMENTATION_ANCHOR)
MAX_SELECTED_PER_CLUSTER = optimizer_int("max_selected_per_cluster", MAX_SELECTED_PER_CLUSTER)
MAX_SELECTED_PER_OPPORTUNITY = optimizer_int("max_selected_per_opportunity", MAX_SELECTED_PER_OPPORTUNITY)
MAX_MODEL_CONVENING_CARDS = optimizer_int("max_model_convening_cards", MAX_MODEL_CONVENING_CARDS)
USE_STATIC_CANDIDATES = optimizer_bool("use_static_candidates", USE_STATIC_CANDIDATES)
STATIC_CANDIDATE_LIMIT = optimizer_int("static_candidate_limit", STATIC_CANDIDATE_LIMIT)


def find_scip_exe() -> Path:
    """Return the SCIP executable used by the local Convening Radar optimizer."""
    candidates = [
        SCIP_EXE,
        SCIP_ROOT / "scip.exe",
        Path(os.getenv("SCIP_EXE", "")) if os.getenv("SCIP_EXE") else None,
    ]
    for candidate in candidates:
        if candidate is not None and candidate.exists():
            return candidate
    raise FileNotFoundError(
        "SCIP executable not found. Expected it at "
        f"{SCIP_EXE}. Set SCIP_EXE to override."
    )

CORE_CONVENING_TOPICS = {
    "versorgung_gesundheit",
    "wirtschaftsentwicklung",
    "kooperationen",
    "technologie",
    "nachhaltigkeit",
    "wissen",
    "pflege",
    "gesundheit",
}

BROAD_MEDIA_MEMBERS = {
    "SWR – Südwestrundfunk",
    "SWR - Südwestrundfunk",
    "ZDF – Zweites Deutsches Fernsehen",
    "ZDF - Zweites Deutsches Fernsehen",
}
BROAD_MEDIA_MEMBERS_NORM = {normalize_search_text(name) for name in BROAD_MEDIA_MEMBERS}

ACADEMIC_MEMBER_MARKERS = (
    "hochschule",
    "technische hochschule",
    "universit",
    "uni ",
    "whu",
)

BUSINESS_MEMBER_MARKERS = (
    " gmbh",
    " ag",
    " se",
    " kg",
    " bank",
    "landesbank",
    "sparkasse",
    "versicherung",
    "werke",
    "holding",
    "industrie",
    "wirtschaft",
)

IMPLEMENTATION_MEMBER_MARKERS = (
    "kammer",
    "verband",
    "klinikum",
    "krankenhaus",
    "stadt ",
    "kreis ",
    "ministerium",
    "agentur",
    "institut",
)

MARKET_CONTEXT_TERMS = {
    "aktien",
    "aktienmarkt",
    "kapitalmarkt",
    "kapitalmärkte",
    "allzeithoch",
    "iwf",
    "ölpreis",
    "oelpreis",
    "waffenruhe",
    "iran",
    "daily",
    "marktkommentar",
    "sentix",
}
MARKET_NOISE_TERMS = {
    "aktien",
    "aktienmarkt",
    "borse",
    "boerse",
    "dax",
    "kapitalmarkt",
    "kapitalmarkte",
    "kursziel",
    "dividende",
    "marktkommentar",
    "sentix",
    "allzeithoch",
}
MARKET_PRESSURE_TERMS = {
    "energiepreise",
    "standortkosten",
    "investitionsbedarf",
    "planungssicherheit",
    "fachkraefte",
    "lieferkette",
    "lieferketten",
    "standort",
    "transformation",
    "wettbewerbsfaehigkeit",
    "versorgungssicherheit",
}
MACRO_TERMS = {
    "schuldenbremse",
    "usa-iran",
    "usa und iran",
    "krieg",
    "iran",
    "aktien",
    "borse",
    "boerse",
    "olpreis",
    "oelpreis",
    "marktkommentar",
    "industrieproduktion",
    "konjunktur",
    "iwf",
    "daily",
}
MARKET_REPORT_TERMS = {
    "arbeitsmarktbericht",
    "arbeitslosenquote",
    "arbeitslosenzahlen",
    "us arbeitsmarkt",
    "us wirtschaft",
    "saisonbereinigt",
    "stellen geschaffen",
    "neue stellen",
}
ACTION_TERMS = {
    "rheinland-pfalz",
    "rlp",
    "kommune",
    "kommunal",
    "kommunen",
    "investition",
    "projekt",
    "kooperation",
    "partnerschaft",
    "programm",
    "infrastruktur",
    "standort",
    "foerderung",
    "umsetzung",
    "pilot",
    "modellprojekt",
    "transfer",
    "qualifizierung",
    "weiterbildung",
}
PROBLEM_TERMS = {
    "versorgung",
    "versorgungslucke",
    "versorgungslucken",
    "praevention",
    "pravention",
    "pflege",
    "fachkraefte",
    "fachkrafte",
    "fachkraft",
    "fachkraftemangel",
    "arbeitsmarkt",
    "kommunen",
    "verwaltung",
    "energiepreise",
    "planungssicherheit",
    "investitionsbedarf",
    "ausbildung",
    "transfer",
    "resilienz",
    "regulierung",
    "genehmigung",
}
PROBLEM_ACTION_VERBS = {
    "fehlt",
    "mangelt",
    "braucht",
    "benoetigt",
    "sichern",
    "umsetzen",
    "skalieren",
    "koordinieren",
    "beschleunigen",
    "ueberfuehren",
    "verbinden",
    "abstimmen",
    "erschliessen",
    "modernisieren",
    "steuern",
    "schuetzen",
    "staerken",
    "starken",
    "erfordert",
    "erfordern",
    "bedarf",
    "umsetzung",
    "kooperation",
    "qualifizierung",
}
VAGUE_PHRASES = {
    "digitalisierung fur nachhaltigkeit, bau oder verwaltung",
    "standort und transformation",
    "gesellschaftliche verantwortung",
    "gemeinsamer problemraum",
    "exploratives treffen",
}
CONCRETE_PROBLEM_TERMS = PROBLEM_TERMS | {
    "ersatzbaustoffverordnung",
    "bauwirtschaft",
    "finanzierung",
    "cybersicherheit",
    "daten",
    "plattform",
    "gesundheitsversorgung",
}

DYNAMIC_FILTER_TERMS_CACHE: dict[str, set[str]] | None = None
DYNAMIC_TERM_CANDIDATES: list[dict[str, Any]] = []

DYNAMIC_PROBLEM_WEIGHT = 0.5
DYNAMIC_ACTION_WEIGHT = 0.5
FEEDBACK_TERM_WEIGHT = 0.8
DYNAMIC_TERM_CONFIDENCE_THRESHOLD = 0.70
DYNAMIC_TERM_EXPIRY_DAYS = 14
DYNAMIC_TERM_PROMOTION_WEEKS = 3
FEEDBACK_TERM_TRUST_THRESHOLD = 2.0
FEEDBACK_TERM_SUPPRESS_THRESHOLD = -2.0
FEEDBACK_CONFIDENCE_STEP = 0.03
FEEDBACK_CONFIDENCE_BOOST_CAP = 0.12
FEEDBACK_CONFIDENCE_PENALTY_CAP = -0.20

CACHE_PROBLEM_HINTS = {
    "bedarf", "mangel", "luecke", "lucke", "engpass", "krise", "problem",
    "risiko", "druck", "herausforderung", "resilienz", "versorgung", "pflege",
    "fachkraft", "arbeitsmarkt", "standort", "investition", "transformation",
    "regulierung", "genehmigung", "sicherheit", "infrastruktur", "digitalisierung",
    "energie", "foerderung",
}

CACHE_ACTION_HINTS = {
    "projekt", "pilot", "modell", "transfer", "kooperation", "partnerschaft",
    "programm", "initiative", "umsetzung", "weiterbildung", "qualifizierung",
    "foerderung", "investition", "modernisierung", "plattform", "testfeld",
    "cloud", "tool", "institut", "labor", "netzwerk", "exchange",
}

CACHE_ACTION_VERB_HINTS = {
    "startet", "entwickelt", "eroeffnet", "eröffnet", "kooperiert", "foerdert",
    "fördert", "testet", "pilotiert", "umsetzen", "setzt um", "baut auf",
    "sichern", "staerken", "stärken", "beschleunigen", "verbinden",
    "koordinieren", "skalieren", "modernisieren", "qualifizieren",
    "investieren", "transformieren", "validieren",
}

CACHE_ASSET_HINTS = {
    "cloud", "plattform", "tool", "institut", "labor", "programm", "netzwerk",
    "exchange", "initiative", "innovationtalk", "library of things", "einbaukarte",
}

CACHE_TERM_STOPLIST = {
    "hochschule", "forschung", "gemeinsam", "aktuelles", "aktuelle", "neu",
    "presse", "meldung", "termin", "veranstaltung", "mainz", "trier",
    "rheinland", "pfalz", "rlp", "gmbh", "universitaet", "universität",
    "digital", "digitale", "digitales", "digitalisierung", "regional",
    "regionale", "entwicklung", "technologie", "wissen", "bildung",
    "soziales", "kultur", "kulturelle", "nachhaltig", "nachhaltige",
    "nachhaltigen", "gesund", "netzwerk",
}


def dynamic_term_context(term: str, events_df: Optional[pd.DataFrame] = None, window: int = 120, limit: int = 12) -> str:
    if events_df is None or events_df.empty:
        return ""
    normalized_term = normalize_search_text(term)
    contexts: list[str] = []
    text_columns = [
        column for column in [
            "titel", "title", "snippet", "mechanism", "implementation_angle",
            "scip_reasoning", "recommended_zirp_use", "briefing_sentence",
            "opportunity_keywords", "hauptsektion", "themenfelder", "mitglied",
        ]
        if column in events_df.columns
    ]
    for _, row in events_df.head(500).iterrows():
        raw = " ".join(str(row.get(column, "") or "") for column in text_columns)
        normalized = normalize_search_text(raw)
        index = normalized.find(normalized_term)
        if index < 0:
            continue
        start = max(0, index - window)
        end = min(len(normalized), index + len(normalized_term) + window)
        contexts.append(normalized[start:end])
        if len(contexts) >= limit:
            break
    return " ".join(contexts)


def classify_dynamic_term_candidate(term: str, context: str = "") -> tuple[str, float, dict[str, int]]:
    normalized = normalize_search_text(term)
    if not normalized or normalized in CACHE_TERM_STOPLIST:
        return "", 0.0, {}
    if len(normalized) < 5 or len(normalized) > 56:
        return "", 0.0, {}
    if not re.search(r"[a-z]", normalized):
        return "", 0.0, {}

    context_text = normalize_search_text(" ".join([normalized, context]))
    problem_hits = normalized_hit_count(CACHE_PROBLEM_HINTS, context_text)
    action_hits = normalized_hit_count(CACHE_ACTION_HINTS | CACHE_ACTION_VERB_HINTS, context_text)
    asset_hits = normalized_hit_count(CACHE_ASSET_HINTS, context_text)
    macro_hits = normalized_hit_count(MACRO_TERMS | MARKET_NOISE_TERMS, context_text)

    scores = {
        "problem": problem_hits * 0.22 + action_hits * 0.05 - macro_hits * 0.12,
        "action": action_hits * 0.22 + asset_hits * 0.06 - macro_hits * 0.10,
        "asset": asset_hits * 0.22 + action_hits * 0.04 - macro_hits * 0.08,
        "noise": macro_hits * 0.25,
    }
    category, raw_score = max(scores.items(), key=lambda item: item[1])
    confidence = max(0.0, min(0.95, raw_score))
    counts = {
        "nearby_problem_words": problem_hits,
        "nearby_action_words": action_hits,
        "nearby_asset_words": asset_hits,
        "nearby_noise_words": macro_hits,
    }
    specific_shape = (
        " " in normalized
        or "-" in normalized
        or len(normalized) >= 13
        or asset_hits >= 1
        or any(marker in normalized for marker in [
            "einbaukarte", "control cloud", "library", "innovationtalk",
            "resilienz", "fachkraft", "versorgung", "infrastruktur",
            "genehmigung", "regulierung", "foerderung", "förderung",
            "investition", "testfeld", "plattform",
        ])
    )
    strong_context = (
        problem_hits >= 2 and action_hits >= 1
    ) or (
        action_hits >= 2 and (problem_hits >= 1 or asset_hits >= 1)
    )
    if category == "noise":
        return category, confidence, counts
    if confidence < DYNAMIC_TERM_CONFIDENCE_THRESHOLD or not specific_shape or not strong_context:
        return "", confidence, counts
    return category, confidence, counts


def parse_iso_date(value: Any) -> Optional[datetime]:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.fromisoformat(text)
    except Exception:
        try:
            return datetime.strptime(text[:10], "%Y-%m-%d")
        except Exception:
            return None


def load_dynamic_term_memory() -> dict[str, dict[str, Any]]:
    if not DYNAMIC_TERM_MEMORY_PATH.exists():
        return {}
    try:
        data = json.loads(DYNAMIC_TERM_MEMORY_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if isinstance(data, list):
        return {
            normalize_search_text(item.get("term", "")): dict(item)
            for item in data
            if isinstance(item, dict) and normalize_search_text(item.get("term", ""))
        }
    if isinstance(data, dict):
        return {
            normalize_search_text(key): dict(value)
            for key, value in data.items()
            if normalize_search_text(key) and isinstance(value, dict)
        }
    return {}


def dynamic_term_memory_status(record: dict[str, Any], now: datetime) -> str:
    status = str(record.get("status", "weekly_dynamic") or "weekly_dynamic")
    if status in {"trusted_feedback", "noise"}:
        return status
    last_seen = parse_iso_date(record.get("last_seen"))
    if last_seen and (now - last_seen).days > int(record.get("expires_after_days", DYNAMIC_TERM_EXPIRY_DAYS) or DYNAMIC_TERM_EXPIRY_DAYS):
        return "inactive"
    return status if status else "weekly_dynamic"


def feedback_score_for_dynamic_terms() -> dict[str, dict[str, Any]]:
    scores: dict[str, dict[str, Any]] = {}
    positive_labels = {"useful"}
    weak_labels = {"interesting_but_weak"}
    negative_labels = {"wrong_connection", "good_topic_wrong_actors", "not_relevant"}
    positive_reasons = {"good_match", "downgrade_not_drop"}
    negative_reasons = {
        "too_macro",
        "insufficient_evidence",
        "source_signal_not_application",
        "missing_rlp_application_anchor",
        "weak_why_now",
        "too_broad_action",
        "wrong_pilot_lane",
        "not_public",
        "source_only_actor",
        "wrong_counterpart",
        "indirect_transfer_legitimacy",
        "good_match_weak_evidence",
        "good_match_weak_rlp",
    }
    target_negative_weight = {
        "evidence_signal": -0.25,
        "rlp_relevance": -0.25,
        "next_action": -0.20,
        "public_framing": -0.20,
    }

    for record in load_feedback_records():
        label = feedback_record_label(record)
        reason = feedback_record_reason(record)
        target = feedback_record_target(record)
        words = record.get("words", [])
        if not isinstance(words, list):
            words = []
        if not words:
            words = normalize_words([], extra_text=" ".join([
                str(record.get("topic", "")),
                str(record.get("reason", "")),
                str(record.get("human_comment", "")),
                str(reason),
            ])) if normalize_words else re.findall(r"[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\-]{2,}", " ".join([
                str(record.get("topic", "")),
                str(record.get("reason", "")),
                str(record.get("human_comment", "")),
                str(reason),
            ]).lower())

        delta = 0.0
        if label in positive_labels:
            delta += 1.0
        elif label in weak_labels:
            delta -= 0.35
        elif label in negative_labels:
            delta -= 1.0

        if reason in positive_reasons:
            delta += 0.6
        if reason in negative_reasons:
            delta -= 0.8
        delta += target_negative_weight.get(target, 0.0)

        selected_value = feedback_selected_value(record)
        if selected_value is False and label in positive_labels:
            delta += 0.35
        if selected_value is True and label in negative_labels:
            delta -= 0.35

        if abs(delta) < 0.01:
            continue

        for raw_word in words[:80]:
            word = normalize_search_text(raw_word)
            if not word or word in CACHE_TERM_STOPLIST or len(word) < 5 or len(word) > 56:
                continue
            item = scores.setdefault(word, {
                "score": 0.0,
                "positive_count": 0,
                "negative_count": 0,
                "labels": Counter(),
                "reasons": Counter(),
            })
            item["score"] += delta
            if delta > 0:
                item["positive_count"] += 1
            else:
                item["negative_count"] += 1
            item["labels"][label or "unknown"] += 1
            item["reasons"][reason or "unspecified"] += 1

    for word, item in scores.items():
        item["score"] = round(float(item.get("score", 0.0)), 3)
        item["labels"] = dict(item["labels"].most_common(5))
        item["reasons"] = dict(item["reasons"].most_common(5))
    return scores


def feedback_dynamic_terms_for(category: str) -> set[str]:
    memory = load_dynamic_term_memory()
    return {
        term
        for term, record in memory.items()
        if dynamic_term_memory_status(record, datetime.now()) == "trusted_feedback"
        and str(record.get("suggested_category", "")) == category
    }


def update_dynamic_term_memory(candidates: list[dict[str, Any]], run_date: Optional[datetime] = None) -> dict[str, dict[str, Any]]:
    now = run_date or datetime.now()
    today = now.strftime("%Y-%m-%d")
    memory = load_dynamic_term_memory()
    feedback_scores = feedback_score_for_dynamic_terms()
    seen_this_run = {
        str(candidate.get("term", "") or "")
        for candidate in candidates
        if int(candidate.get("activated", 0) or 0) == 1 and str(candidate.get("term", "") or "")
    }

    for term, record in list(memory.items()):
        status = dynamic_term_memory_status(record, now)
        if status != record.get("status"):
            record["status"] = status
        if term not in seen_this_run and status == "weekly_dynamic":
            last_seen = parse_iso_date(record.get("last_seen"))
            if last_seen and (now - last_seen).days > int(record.get("expires_after_days", DYNAMIC_TERM_EXPIRY_DAYS) or DYNAMIC_TERM_EXPIRY_DAYS):
                record["status"] = "inactive"
        memory[term] = record

    for candidate in candidates:
        term = str(candidate.get("term", "") or "")
        if not term:
            continue
        candidate_feedback_score = float(candidate.get("human_feedback_score", 0) or 0)
        if int(candidate.get("activated", 0) or 0) != 1 and abs(candidate_feedback_score) < 0.01:
            continue
        record = memory.get(term, {})
        first_seen = record.get("first_seen") or today
        previous_last_seen = record.get("last_seen")
        weeks_seen = int(record.get("weeks_seen", 0) or 0)
        if previous_last_seen != today:
            weeks_seen += 1
        status = dynamic_term_memory_status(record, now) if record else "weekly_dynamic"
        if status == "inactive":
            status = "weekly_dynamic"
        feedback_meta = feedback_scores.get(term, {})
        if feedback_meta:
            human_feedback_score = float(feedback_meta.get("score", 0) or 0)
        else:
            human_feedback_score = float(record.get("human_feedback_score", 0) or 0)
        human_feedback_score = max(-8.0, min(8.0, human_feedback_score))
        suggested_category = str(candidate.get("suggested_category", record.get("suggested_category", "")) or "")
        promotable_category = suggested_category in {"problem", "action", "asset"}
        if human_feedback_score <= FEEDBACK_TERM_SUPPRESS_THRESHOLD:
            status = "noise"
        elif status == "trusted_feedback" and not promotable_category:
            status = "weekly_dynamic"
        elif status == "weekly_dynamic" and promotable_category and (
            human_feedback_score >= FEEDBACK_TERM_TRUST_THRESHOLD
            or (weeks_seen >= DYNAMIC_TERM_PROMOTION_WEEKS and human_feedback_score > 0)
        ):
            status = "trusted_feedback"
        if candidate.get("suggested_category") == "noise" and float(candidate.get("confidence", 0) or 0) >= DYNAMIC_TERM_CONFIDENCE_THRESHOLD:
            status = "noise"
        memory[term] = {
            **record,
            "term": term,
            "suggested_category": suggested_category,
            "first_seen": first_seen,
            "last_seen": today,
            "weeks_seen": weeks_seen,
            "status": status,
            "expires_after_days": int(record.get("expires_after_days", DYNAMIC_TERM_EXPIRY_DAYS) or DYNAMIC_TERM_EXPIRY_DAYS),
            "last_confidence": float(candidate.get("confidence", 0) or 0),
            "last_weighted_score": float(candidate.get("weighted_score", 0) or 0),
            "human_feedback_score": human_feedback_score,
            "feedback_positive_count": int(feedback_meta.get("positive_count", record.get("feedback_positive_count", 0)) or 0),
            "feedback_negative_count": int(feedback_meta.get("negative_count", record.get("feedback_negative_count", 0)) or 0),
            "feedback_labels": feedback_meta.get("labels", record.get("feedback_labels", {})),
            "feedback_reasons": feedback_meta.get("reasons", record.get("feedback_reasons", {})),
        }

    DYNAMIC_TERM_MEMORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    ordered = dict(sorted(memory.items(), key=lambda item: item[0]))
    DYNAMIC_TERM_MEMORY_PATH.write_text(json.dumps(ordered, ensure_ascii=False, indent=2), encoding="utf-8")
    return ordered


def dynamic_term_is_active(term: str, category: str, memory: dict[str, dict[str, Any]], now: datetime) -> bool:
    record = memory.get(term, {})
    status = dynamic_term_memory_status(record, now) if record else "weekly_dynamic"
    if status == "noise" or status == "inactive":
        return False
    if status == "trusted_feedback":
        return True
    return category in {"problem", "action", "asset"}



def load_dynamic_filter_terms(
    weighted_path: Optional[Path] = None,
    events_df: Optional[pd.DataFrame] = None,
) -> dict[str, set[str]]:
    global DYNAMIC_FILTER_TERMS_CACHE, DYNAMIC_TERM_CANDIDATES
    if weighted_path is None and events_df is None and DYNAMIC_FILTER_TERMS_CACHE is not None:
        return DYNAMIC_FILTER_TERMS_CACHE

    terms = {"problem": set(), "action": set(), "asset": set(), "noise": set()}
    candidates: list[dict[str, Any]] = []
    now = datetime.now()
    memory = load_dynamic_term_memory()
    feedback_scores = feedback_score_for_dynamic_terms()
    try:
        path = weighted_path or latest_file("zirp_begriffe_gewichtet")
        df = read_csv(path)
    except Exception:
        if weighted_path is None:
            DYNAMIC_FILTER_TERMS_CACHE = terms
            DYNAMIC_TERM_CANDIDATES = candidates
        return terms

    if "begriff" not in df.columns:
        if weighted_path is None:
            DYNAMIC_FILTER_TERMS_CACHE = terms
            DYNAMIC_TERM_CANDIDATES = candidates
        return terms

    df = df.copy()
    if "weighted_score" in df.columns:
        df["_score"] = pd.to_numeric(df["weighted_score"], errors="coerce").fillna(0)
        df = df.sort_values("_score", ascending=False).head(220)
    else:
        df["_score"] = 0
        df = df.head(220)

    for _, row in df.iterrows():
        raw_term = str(row.get("begriff", "") or "")
        normalized = normalize_search_text(raw_term)
        context = dynamic_term_context(normalized, events_df)
        category, confidence, counts = classify_dynamic_term_candidate(normalized, context)
        feedback_meta = feedback_scores.get(normalized, {})
        feedback_score = float(feedback_meta.get("score", 0) or 0)
        if feedback_score:
            confidence = max(0.0, min(0.95, confidence + max(FEEDBACK_CONFIDENCE_PENALTY_CAP, min(FEEDBACK_CONFIDENCE_BOOST_CAP, feedback_score * FEEDBACK_CONFIDENCE_STEP))))
        memory_record = memory.get(normalized, {})
        memory_status = dynamic_term_memory_status(memory_record, now) if memory_record else "new"
        if feedback_score <= FEEDBACK_TERM_SUPPRESS_THRESHOLD:
            memory_status = "noise"
        elif feedback_score >= FEEDBACK_TERM_TRUST_THRESHOLD and category:
            memory_status = "trusted_feedback"
        activated = bool(
            category
            and confidence >= DYNAMIC_TERM_CONFIDENCE_THRESHOLD
            and feedback_score > FEEDBACK_TERM_SUPPRESS_THRESHOLD
            and dynamic_term_is_active(normalized, category, memory, now)
        )
        record = {
            "term": normalized,
            "suggested_category": category or "review",
            "frequency": int(float(row.get("raw_count", 0) or 0)),
            "weighted_score": float(row.get("_score", 0) or 0),
            "nearby_verbs": counts.get("nearby_action_words", 0),
            "nearby_problem_words": counts.get("nearby_problem_words", 0),
            "nearby_action_words": counts.get("nearby_action_words", 0),
            "nearby_asset_words": counts.get("nearby_asset_words", 0),
            "nearby_noise_words": counts.get("nearby_noise_words", 0),
            "appears_in_selected_cards": 0,
            "appears_in_rejected_cards": 0,
            "human_feedback_score": float(memory_record.get("human_feedback_score", 0) or 0) + feedback_score,
            "feedback_positive_count": int(feedback_meta.get("positive_count", 0) or 0),
            "feedback_negative_count": int(feedback_meta.get("negative_count", 0) or 0),
            "confidence": round(confidence, 3),
            "memory_status": memory_status,
            "first_seen": memory_record.get("first_seen", ""),
            "last_seen": memory_record.get("last_seen", ""),
            "weeks_seen": int(memory_record.get("weeks_seen", 0) or 0),
            "expires_after_days": int(memory_record.get("expires_after_days", DYNAMIC_TERM_EXPIRY_DAYS) or DYNAMIC_TERM_EXPIRY_DAYS),
            "activated": int(activated),
        }
        candidates.append(record)
        if activated:
            terms.setdefault(category, set()).add(normalized)

    memory = update_dynamic_term_memory(candidates, now)
    for candidate in candidates:
        term = str(candidate.get("term", "") or "")
        if not term:
            continue
        record = memory.get(term, {})
        candidate["memory_status"] = dynamic_term_memory_status(record, now) if record else candidate.get("memory_status", "new")
        candidate["first_seen"] = record.get("first_seen", candidate.get("first_seen", ""))
        candidate["last_seen"] = record.get("last_seen", candidate.get("last_seen", ""))
        candidate["weeks_seen"] = int(record.get("weeks_seen", candidate.get("weeks_seen", 0)) or 0)
        candidate["expires_after_days"] = int(record.get("expires_after_days", candidate.get("expires_after_days", DYNAMIC_TERM_EXPIRY_DAYS)) or DYNAMIC_TERM_EXPIRY_DAYS)
        candidate["human_feedback_score"] = float(record.get("human_feedback_score", candidate.get("human_feedback_score", 0)) or 0)
        candidate["feedback_positive_count"] = int(record.get("feedback_positive_count", candidate.get("feedback_positive_count", 0)) or 0)
        candidate["feedback_negative_count"] = int(record.get("feedback_negative_count", candidate.get("feedback_negative_count", 0)) or 0)

    if weighted_path is None:
        DYNAMIC_FILTER_TERMS_CACHE = terms
        DYNAMIC_TERM_CANDIDATES = candidates
    else:
        DYNAMIC_FILTER_TERMS_CACHE = terms
        DYNAMIC_TERM_CANDIDATES = candidates
    return terms


def dynamic_terms_for(category: str) -> set[str]:
    return load_dynamic_filter_terms().get(category, set())


def weighted_problem_hit_score(text: str) -> dict[str, float]:
    normalized = normalize_search_text(text)
    dynamic = load_dynamic_filter_terms()
    feedback_terms = feedback_dynamic_terms_for("problem")
    static_problem_hits = normalized_hit_count(CONCRETE_PROBLEM_TERMS, normalized)
    dynamic_problem_hits = normalized_hit_count(dynamic.get("problem", set()) - feedback_terms, normalized)
    feedback_problem_hits = normalized_hit_count(feedback_terms, normalized)
    return {
        "static": float(static_problem_hits),
        "dynamic": float(dynamic_problem_hits),
        "feedback": float(feedback_problem_hits),
        "score": (
            static_problem_hits
            + dynamic_problem_hits * DYNAMIC_PROBLEM_WEIGHT
            + feedback_problem_hits * FEEDBACK_TERM_WEIGHT
        ),
    }


def weighted_action_hit_score(text: str, *, verbs: bool = False) -> dict[str, float]:
    normalized = normalize_search_text(text)
    dynamic = load_dynamic_filter_terms()
    static_terms = PROBLEM_ACTION_VERBS if verbs else ACTION_TERMS
    feedback_terms = feedback_dynamic_terms_for("action")
    dynamic_terms = dynamic.get("action", set()) - feedback_terms
    static_action_hits = normalized_hit_count(static_terms, normalized)
    dynamic_action_hits = normalized_hit_count(dynamic_terms, normalized)
    feedback_action_hits = normalized_hit_count(feedback_terms, normalized)
    return {
        "static": float(static_action_hits),
        "dynamic": float(dynamic_action_hits),
        "feedback": float(feedback_action_hits),
        "score": (
            static_action_hits
            + dynamic_action_hits * DYNAMIC_ACTION_WEIGHT
            + feedback_action_hits * FEEDBACK_TERM_WEIGHT
        ),
    }


def export_dynamic_term_candidates(path: Path, patterns_df: Optional[pd.DataFrame] = None) -> None:
    rows = [dict(row) for row in DYNAMIC_TERM_CANDIDATES]
    if patterns_df is not None and rows and not patterns_df.empty:
        selected_text = normalize_search_text(" ".join(
            " ".join(str(row.get(column, "") or "") for column in [
                "members", "convening_theme", "shared_topics", "bridge_topics",
                "concrete_clusters", "recent_signals", "decision_summary",
            ])
            for _, row in patterns_df[patterns_df.get("selected", 0).astype(int) == 1].iterrows()
        ))
        rejected_text = normalize_search_text(" ".join(
            " ".join(str(row.get(column, "") or "") for column in [
                "members", "convening_theme", "shared_topics", "bridge_topics",
                "concrete_clusters", "recent_signals", "decision_summary",
            ])
            for _, row in patterns_df[patterns_df.get("selected", 0).astype(int) != 1].iterrows()
        ))
        for row in rows:
            term = str(row.get("term", ""))
            row["appears_in_selected_cards"] = int(term in selected_text) if term else 0
            row["appears_in_rejected_cards"] = int(term in rejected_text) if term else 0
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")


THEME_WEIGHTS = {
    "versorgung_gesundheit": 6.0,
    "wirtschaftsentwicklung": 5.5,
    "kooperationen": 5.0,
    "fuehrungswechsel": 4.0,
    "soziales_engagement": 3.5,
    "gesundheit": 4.0,
    "pflege": 4.0,
    "wirtschaft": 3.8,
    "technologie": 3.8,
    "nachhaltigkeit": 3.5,
    "wissen": 3.2,
    "gesellschaft": 3.0,
    "kultur": 2.0,
}

COMPLEMENTARY_TOPIC_PAIRS = {
    frozenset(("wirtschaft", "technologie")): 4.0,
    frozenset(("wirtschaft", "nachhaltigkeit")): 4.0,
    frozenset(("versorgung_gesundheit", "wissen")): 4.0,
    frozenset(("versorgung_gesundheit", "gesellschaft")): 3.0,
    frozenset(("kooperationen", "wissen")): 3.0,
    frozenset(("kooperationen", "technologie")): 3.0,
    frozenset(("technologie", "nachhaltigkeit")): 3.0,
}

ZIRP_CLUSTER_TOPICS = {
    "nachhaltigkeit": {
        "energie_versorgungssicherheit": {
            "label": "Energie, Netze und Versorgungssicherheit",
            "problem": "Wie bleiben Unternehmen, Kommunen und Infrastrukturträger in Rheinland-Pfalz investitions- und versorgungssicher?",
            "signals": ["energiepreise", "strompreise", "wärme", "waerme", "netzausbau", "versorgungssicherheit", "erneuerbare energien", "speicher", "photovoltaik", "wasserstoff", "standortkosten", "energieintensive industrie"],
        },
        "kreislaufwirtschaft_bau_ebv": {
            "label": "Kreislaufwirtschaft, Bau und Ersatzbaustoffverordnung",
            "problem": "Wie können Bauwirtschaft, Rohstoffindustrie, Kommunen und Verwaltung Kreislaufwirtschaft praktisch und rechtssicher umsetzen?",
            "signals": ["ersatzbaustoffverordnung", "ebv", "einbaukarte", "recyclingbaustoffe", "rohstoffindustrie", "bauwirtschaft", "kommunen", "verwaltung", "planungssicherheit", "regulierung", "genehmigung"],
        },
        "wasser_umwelttechnik_resilienz": {
            "label": "Wasser, Umwelttechnik und kommunale Resilienz",
            "problem": "Wie können Kommunen und Infrastrukturträger Wasser-, Abwasser- und Umwelttechnik digitaler und resilienter steuern?",
            "signals": ["wasser", "abwasser", "umwelttechnik", "prozesssteuerung", "sensorik", "cloud", "kommunale infrastruktur", "klimafolgen", "resilienz", "versorgung", "entsorgung"],
        },
        "klima_anpassung_stadt_land": {
            "label": "Klimaanpassung in Stadt und Land",
            "problem": "Wie reagieren Kommunen, Bauwirtschaft, Landwirtschaft und Infrastruktur auf Hitze, Starkregen und Flächendruck?",
            "signals": ["starkregen", "hitze", "klimaanpassung", "flächenverbrauch", "flaechenverbrauch", "innenstadt", "mobilität", "mobilitaet", "stadtentwicklung", "biodiversität", "biodiversitaet", "landwirtschaft", "kommunen"],
        },
    },
    "wirtschaft": {
        "standort_investition_transformation": {
            "label": "Standort, Investition und industrielle Transformation",
            "problem": "Welche Standortbedingungen entscheiden darüber, ob Unternehmen in Rheinland-Pfalz investieren, modernisieren oder verlagern?",
            "signals": ["investition", "standort", "produktion", "werk", "kapazität", "kapazitaet", "arbeitsplätze", "arbeitsplaetze", "modernisierung", "transformation", "wettbewerbsfähigkeit", "wettbewerbsfaehigkeit", "industrie", "export", "lieferketten"],
        },
        "mittelstand_digitalisierung_produktivitaet": {
            "label": "Mittelstand, Digitalisierung und Produktivität",
            "problem": "Wie können mittelständische Unternehmen digitale Werkzeuge praktisch einsetzen, ohne dass daraus nur Pilotprojekte entstehen?",
            "signals": ["digitalisierung", "mittelstand", "produktivität", "produktivitaet", "automatisierung", "ki", "software", "plattform", "prozessoptimierung", "industrie 4.0", "cloud", "erp"],
        },
        "arbeitsmarkt_fachkraefte": {
            "label": "Arbeitsmarkt, Fachkräfte und Qualifizierung",
            "problem": "Wie sichern Unternehmen, Hochschulen, Kammern und Arbeitsverwaltung Fachkräfte für Transformation, Pflege, Technik und Verwaltung?",
            "signals": ["fachkräfte", "fachkraefte", "ausbildung", "weiterbildung", "qualifizierung", "duale ausbildung", "arbeitsmarkt", "personal", "recruiting", "integration", "zuwanderung", "nachwuchs"],
        },
        "finanzierung_foerderung_transformation": {
            "label": "Finanzierung und Förderung von Transformation",
            "problem": "Wie kommen Investitionen, Fördermittel, Banken und Unternehmensstrategien schneller zusammen?",
            "signals": ["förderung", "foerderung", "finanzierung", "kredit", "investition", "kapital", "transformation", "mittelstand", "nachfolge", "innovation", "risiko", "bank"],
        },
    },
    "technologie": {
        "ki_daten_verwaltung_wirtschaft": {
            "label": "KI, Daten und Verwaltungs-/Wirtschaftsanwendungen",
            "problem": "Wo können KI und Datenplattformen in Unternehmen, Verwaltung und öffentlicher Infrastruktur konkret Nutzen stiften?",
            "signals": ["ki", "künstliche intelligenz", "kuenstliche intelligenz", "daten", "plattform", "automatisierung", "verwaltung", "e-government", "prozessdigitalisierung", "chatbot", "wissensmanagement"],
        },
        "industrie_40_automation": {
            "label": "Industrie 4.0, Automatisierung und Produktion",
            "problem": "Wie wird industrielle Produktion in Rheinland-Pfalz digitaler, robuster und wettbewerbsfähiger?",
            "signals": ["industrie 4.0", "automatisierung", "produktion", "sensorik", "maschinenbau", "digitaler zwilling", "robotik", "prozesssteuerung", "fertigung", "smart factory"],
        },
        "cyber_resilienz_kritische_infrastruktur": {
            "label": "Cyberresilienz und kritische Infrastruktur",
            "problem": "Wie schützen Unternehmen, Kommunen und Versorger digitale Systeme, Daten und kritische Infrastruktur?",
            "signals": ["cybersicherheit", "it-sicherheit", "resilienz", "kritische infrastruktur", "krisenmanagement", "datenschutz", "cloud", "netzwerke", "kommunen", "versorgung"],
        },
        "digitale_bau_verwaltung_planung": {
            "label": "Digitale Bau-, Planungs- und Genehmigungsprozesse",
            "problem": "Wie können Bau, Planung, Genehmigung und Verwaltung durch digitale Werkzeuge schneller und rechtssicherer werden?",
            "signals": ["bim", "planung", "genehmigung", "bauverwaltung", "einbaukarte", "ersatzbaustoffverordnung", "kommunen", "verwaltung", "digitales tool", "plattform"],
        },
    },
    "kultur": {
        "kultur_standort_identitaet": {
            "label": "Kultur, Standortidentität und regionale Attraktivität",
            "problem": "Wie stärkt Kultur die Attraktivität von Städten, Regionen und Arbeitgeberstandorten?",
            "signals": ["kultur", "innenstadt", "stadtentwicklung", "regionalität", "regionalitaet", "identität", "identitaet", "tourismus", "arbeitgeberattraktivität", "arbeitgeberattraktivitaet", "standortmarketing", "veranstaltung", "festival"],
        },
        "kreativwirtschaft_digitale_medien": {
            "label": "Kreativwirtschaft, Medien und digitale Öffentlichkeit",
            "problem": "Wie verändern Medien, Plattformen und kreative Arbeit die öffentliche Wahrnehmung von Rheinland-Pfalz?",
            "signals": ["medien", "kreativwirtschaft", "film", "audio", "plattform", "publikation", "digitale öffentlichkeit", "digitale oeffentlichkeit", "kommunikation", "journalismus", "kulturvermittlung"],
        },
        "musikstipendium_talentfoerderung": {
            "label": "Talentförderung, Musik und Nachwuchs",
            "problem": "Wie lassen sich kulturelle Talentförderung, Bildung und regionale Sichtbarkeit verbinden?",
            "signals": ["musikstipendium", "stipendium", "nachwuchs", "talent", "förderung", "foerderung", "bildung", "musik", "hochschule", "kultur"],
        },
    },
    "wissen": {
        "hochschule_wirtschaft_transfer": {
            "label": "Hochschule-Wirtschaft-Transfer",
            "problem": "Wie kommen Forschung, Unternehmen und öffentliche Akteure schneller in gemeinsame Umsetzungsprojekte?",
            "signals": ["transfer", "kooperation", "forschung", "hochschule", "unternehmen", "projekt", "innovation", "anwendung", "gründung", "gruendung", "wissensallianz"],
        },
        "weiterbildung_lebenslanges_lernen": {
            "label": "Weiterbildung und lebenslanges Lernen",
            "problem": "Wie können Beschäftigte, Verwaltung und Unternehmen Weiterbildung in reale Transformationsprozesse integrieren?",
            "signals": ["weiterbildung", "qualifizierung", "zertifikat", "lebenslanges lernen", "fachkräfte", "fachkraefte", "kompetenzen", "digitalisierung", "pflege", "verwaltung", "technik"],
        },
        "wissenschaft_demokratie_gesellschaft": {
            "label": "Wissenschaft, Demokratie und gesellschaftlicher Zusammenhalt",
            "problem": "Wie kann wissenschaftliches Wissen in gesellschaftliche Debatten, Demokratiebildung und öffentliche Orientierung wirken?",
            "signals": ["demokratie", "gesellschaft", "wissenschaftskommunikation", "bildung", "teilhabe", "extremismus", "medienkompetenz", "zivilgesellschaft", "forschung", "öffentlichkeit", "oeffentlichkeit"],
        },
        "verwaltungswissen_digitalisierung": {
            "label": "Verwaltungswissen und digitale Umsetzung",
            "problem": "Wie wird Verwaltung bei Regulierung, Genehmigung, Digitalisierung und Umsetzung leistungsfähiger?",
            "signals": ["verwaltung", "e-government", "genehmigung", "regulierung", "kommunen", "digitalisierung", "umsetzung", "planung", "prozess", "daten"],
        },
    },
    "gesellschaft": {
        "gesundheit_versorgung_praevention": {
            "label": "Gesundheitsversorgung und Prävention",
            "problem": "Wie kann Gesundheitsversorgung in Rheinland-Pfalz zugänglich, präventiv und digital unterstützt werden?",
            "signals": ["gesundheit", "versorgung", "prävention", "praevention", "patienten", "krankenkasse", "ärzte", "aerzte", "notfallversorgung", "digital health", "e-health", "ambulant"],
        },
        "pflege_fachkraefte_sozialwirtschaft": {
            "label": "Pflege, Fachkräfte und Sozialwirtschaft",
            "problem": "Wie sichern Pflege- und Sozialträger Personal, Qualität und Versorgung in einer alternden Gesellschaft?",
            "signals": ["pflege", "fachkräfte", "fachkraefte", "sozialwirtschaft", "versorgung", "teilhabe", "ausbildung", "personal", "gesundheitsberufe", "inklusion", "betreuung"],
        },
        "teilhabe_inklusion_arbeitsmarkt": {
            "label": "Teilhabe, Inklusion und Arbeitsmarkt",
            "problem": "Wie können Menschen mit Familienpflichten, Behinderung oder schwierigen Erwerbsbiografien besser am Arbeitsmarkt teilnehmen?",
            "signals": ["teilhabe", "inklusion", "arbeitsmarkt", "familie", "mütter", "muetter", "vereinbarkeit", "integration", "barrierefreiheit", "soziale innovation", "beschäftigung", "beschaeftigung"],
        },
        "demokratie_medien_zusammenhalt": {
            "label": "Demokratie, Medien und gesellschaftlicher Zusammenhalt",
            "problem": "Wie stärken Medien, Bildung, Kirchen, Wissenschaft und Zivilgesellschaft Vertrauen und demokratische Orientierung?",
            "signals": ["demokratie", "zusammenhalt", "medien", "gesellschaft", "bildung", "vielfalt", "extremismus", "dialog", "kirche", "öffentlichkeit", "oeffentlichkeit"],
        },
        "sport_region_identifikation": {
            "label": "Sport, Region und gesellschaftliche Identifikation",
            "problem": "Wie können Sportvereine als regionale Plattformen für Bildung, Teilhabe, Gesundheit oder Standortbindung wirken?",
            "signals": ["sport", "verein", "nachwuchs", "bildung", "gesundheit", "integration", "fans", "region", "identifikation", "gesellschaftliches engagement"],
        },
    },
}

CLUSTER_TOPIC_LOOKUP = {
    cluster_id: {"field": field, **cluster}
    for field, clusters in ZIRP_CLUSTER_TOPICS.items()
    for cluster_id, cluster in clusters.items()
}
CLUSTER_TOPIC_IDS = set(CLUSTER_TOPIC_LOOKUP)
CLUSTER_TOPIC_ALIASES = {f"cluster:{cluster_id}" for cluster_id in CLUSTER_TOPIC_IDS}
ZIRP_FIELDS = set(ZIRP_CLUSTER_TOPICS)
BROAD_FIELD_TOPICS = ZIRP_FIELDS | {
    "technologie", "gesellschaft", "wissen", "wirtschaft", "nachhaltigkeit", "kultur",
    "kooperationen", "soziales_engagement", "wirtschaftsentwicklung", "versorgung_gesundheit",
}
CLUSTER_NOISE_TERMS = {
    "geschäftsführer vorgestellt", "geschaeftsfuehrer vorgestellt", "vizepräsident", "vizepraesident",
    "ranking", "che-ranking", "bestnoten", "podcast", "folge", "ticket", "sponsoring",
    "gewinnspiel", "infoabend", "alumni", "messeauftritt", "dissertation",
}
PRIORITY_CLUSTER_WEIGHTS = {
    "kreislaufwirtschaft_bau_ebv": 9.0,
    "energie_versorgungssicherheit": 8.8,
    "arbeitsmarkt_fachkraefte": 8.6,
    "gesundheit_versorgung_praevention": 8.4,
    "pflege_fachkraefte_sozialwirtschaft": 8.3,
    "hochschule_wirtschaft_transfer": 8.1,
    "ki_daten_verwaltung_wirtschaft": 7.9,
    "wasser_umwelttechnik_resilienz": 7.7,
    "standort_investition_transformation": 7.5,
    "digitale_bau_verwaltung_planung": 7.4,
    "demokratie_medien_zusammenhalt": 7.0,
}

CLUSTER_MEMBER_RELEVANCE = {
    "energie_versorgungssicherheit": {"Pfalzwerke AG": 5, "Westenergie AG": 5, "ENTEGA Plus GmbH": 5, "Mainzer Stadtwerke AG": 5, "JUWI GmbH": 5, "Schoenergie GmbH": 4, "KÜBLER GmbH": 4, "BASF SE": 4, "SCHOTT AG": 4, "thyssenkrupp Rasselstein GmbH": 4, "SIMONA AG": 3, "Michelin Reifenwerke AG & Co. KGaA": 3, "Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 3, "Landesbank Baden-Württemberg (LBBW)": 3, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 3, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 3, "Technische Hochschule Bingen": 3, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 3},
    "kreislaufwirtschaft_bau_ebv": {"vero – Verband der Bau- und Rohstoffindustrie": 5, "Bauwirtschaft Rheinland-Pfalz e.V.": 5, "Architektenkammer Rheinland-Pfalz": 4, "Ingenieurkammer Rheinland-Pfalz": 4, "Städtetag Rheinland-Pfalz": 5, "Landkreistag Rheinland-Pfalz": 5, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 3, "Karl Gemünden GmbH & Co. KG": 4, "Heberger Gruppe": 4, "Wilhelm Faber GmbH": 3, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "Deutsche Universität für Verwaltungswissenschaften Speyer": 3, "Hochschule Mainz": 3, "Technische Hochschule Bingen": 3, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 3, "Entwicklungsagentur Rheinland-Pfalz e.V.": 3, "Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 3},
    "wasser_umwelttechnik_resilienz": {"Zahnen Technik GmbH": 5, "KSB AG": 5, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "Technische Hochschule Bingen": 4, "Hochschule Trier": 4, "Hochschule Koblenz": 3, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 4, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Mainzer Stadtwerke AG": 4, "Pfalzwerke AG": 3, "Westenergie AG": 3, "ENTEGA Plus GmbH": 3, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 3, "Staatskanzlei des Landes Rheinland-Pfalz": 3, "Entwicklungsagentur Rheinland-Pfalz e.V.": 3, "BASF SE": 3, "TH Bingen": 4},
    "klima_anpassung_stadt_land": {"Städtetag Rheinland-Pfalz": 5, "Landkreistag Rheinland-Pfalz": 5, "Entwicklungsagentur Rheinland-Pfalz e.V.": 5, "Bauern- und Winzerverband Rheinland-Pfalz Süd e.V.": 5, "Bauwirtschaft Rheinland-Pfalz e.V.": 4, "Architektenkammer Rheinland-Pfalz": 4, "Ingenieurkammer Rheinland-Pfalz": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 3, "Technische Hochschule Bingen": 4, "Hochschule Mainz": 3, "Hochschule Trier": 3, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 4, "Johannes Gutenberg-Universität Mainz": 3, "Pfalzwerke AG": 3, "Westenergie AG": 3, "Mainzer Stadtwerke AG": 4, "Transdev GmbH": 3, "JUWI GmbH": 3},
    "standort_investition_transformation": {"BASF SE": 5, "Boehringer Ingelheim Pharma GmbH & Co. KG": 5, "SCHOTT AG": 5, "thyssenkrupp Rasselstein GmbH": 5, "Michelin Reifenwerke AG & Co. KGaA": 4, "KSB AG": 4, "Joseph Vögele AG": 4, "Continental Automotive Technologies GmbH, Safety and Motion (SAM), Standort Rheinböllen": 4, "SIMONA AG": 4, "Bitburger Braugruppe GmbH": 4, "Eckes-Granini Deutschland GmbH": 4, "GLOBUS Markthallen Holding GmbH & Co. KG": 4, "Debeka Versicherungsgruppe": 4, "Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 5, "Landesbank Baden-Württemberg (LBBW)": 5, "Sparkassenverband Rheinland-Pfalz": 4, "Vereinigte VR Bank Kur- und Rheinpfalz eG": 3, "KPMG AG Wirtschaftsprüfungsgesellschaft": 4, "Ernst & Young GmbH": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 5, "Staatskanzlei des Landes Rheinland-Pfalz": 4, "WHU – Otto Beisheim School of Management": 4},
    "mittelstand_digitalisierung_produktivitaet": {"Empolis Information Management GmbH": 5, "BRICKMAKERS AG": 5, "ITK Engineering GmbH": 5, "SmartFactory KL e.V.": 5, "robotspaceship": 5, "Aturas GmbH": 4, "IQIB - Institut für qualifizierende Innovationsforschung und -beratung GmbH": 4, "Deutsche Telekom AG": 5, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "Handwerkskammern Rheinland-Pfalz": 4, "Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 3, "Sparkassenverband Rheinland-Pfalz": 3, "WHU – Otto Beisheim School of Management": 4, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 5, "Hochschule Kaiserslautern": 4, "Hochschule Mainz": 4, "Hochschule Koblenz": 3, "Hochschule Trier": 3, "Technische Hochschule Bingen": 3, "PFAFF Industriesysteme und Maschinen GmbH": 4, "Joseph Vögele AG": 4, "KÜBLER GmbH": 3, "Licharz GmbH": 3},
    "arbeitsmarkt_fachkraefte": {"Bundesagentur für Arbeit, Regionaldirektion Rheinland-Pfalz-Saarland": 5, "Duale Hochschule Rheinland-Pfalz": 5, "Handwerkskammern Rheinland-Pfalz": 5, "Deutscher Gewerkschaftsbund (DGB) Rheinland-Pfalz/Saarland": 5, "JOBS FOR MOMS UG": 4, "Leonardo PersonalKonzept GmbH": 4, "Hochschule für Wirtschaft und Gesellschaft Ludwigshafen": 4, "Hochschule Kaiserslautern": 4, "Hochschule Koblenz": 4, "Hochschule Mainz": 4, "Hochschule Trier": 4, "Hochschule Worms": 4, "Technische Hochschule Bingen": 4, "Katholische Hochschule Mainz": 4, "Vinzenz Pallotti University gGmbH": 3, "WHU – Otto Beisheim School of Management": 4, "Caritasverband für die Diözese Speyer e.V.": 4, "Alexianer GmbH": 4, "Barmherzige Brüder Trier gGmbH": 4, "in.betrieb gGmbH": 4, "GLOBUS Markthallen Holding GmbH & Co. KG": 3, "BASF SE": 3, "Boehringer Ingelheim Pharma GmbH & Co. KG": 3, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4},
    "finanzierung_foerderung_transformation": {"Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 5, "Landesbank Baden-Württemberg (LBBW)": 5, "Sparkassenverband Rheinland-Pfalz": 5, "Vereinigte VR Bank Kur- und Rheinpfalz eG": 4, "KPMG AG Wirtschaftsprüfungsgesellschaft": 4, "Ernst & Young GmbH": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 5, "Staatskanzlei des Landes Rheinland-Pfalz": 4, "BASF SE": 3, "Boehringer Ingelheim Pharma GmbH & Co. KG": 3, "SCHOTT AG": 3, "JUWI GmbH": 3, "Pfalzwerke AG": 3, "Westenergie AG": 3, "Debeka Versicherungsgruppe": 4, "Provinzial Versicherung AG": 4, "WHU – Otto Beisheim School of Management": 4, "Handwerkskammern Rheinland-Pfalz": 3},
    "ki_daten_verwaltung_wirtschaft": {"Empolis Information Management GmbH": 5, "BRICKMAKERS AG": 5, "robotspaceship": 5, "Aturas GmbH": 4, "Deutsche Telekom AG": 5, "Deutsche Universität für Verwaltungswissenschaften Speyer": 5, "Universität Trier": 4, "Universität Koblenz": 4, "Johannes Gutenberg-Universität Mainz": 4, "Hochschule Mainz": 4, "Hochschule Kaiserslautern": 4, "Hochschule Koblenz": 3, "Hochschule Trier": 3, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 5, "Staatskanzlei des Landes Rheinland-Pfalz": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "SmartFactory KL e.V.": 4, "ITK Engineering GmbH": 4, "IQIB - Institut für qualifizierende Innovationsforschung und -beratung GmbH": 4},
    "industrie_40_automation": {"SmartFactory KL e.V.": 5, "ITK Engineering GmbH": 5, "Joseph Vögele AG": 5, "PFAFF Industriesysteme und Maschinen GmbH": 5, "Continental Automotive Technologies GmbH, Safety and Motion (SAM), Standort Rheinböllen": 5, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 5, "Hochschule Kaiserslautern": 4, "Technische Hochschule Bingen": 4, "Hochschule Trier": 3, "SCHOTT AG": 4, "BASF SE": 4, "KSB AG": 4, "thyssenkrupp Rasselstein GmbH": 4, "Michelin Reifenwerke AG & Co. KGaA": 4, "SIMONA AG": 3, "KÜBLER GmbH": 3, "Licharz GmbH": 3, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "Deutsche Telekom AG": 4, "Empolis Information Management GmbH": 4},
    "cyber_resilienz_kritische_infrastruktur": {"Deutsche Telekom AG": 5, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 5, "Pfalzwerke AG": 5, "Westenergie AG": 5, "ENTEGA Plus GmbH": 5, "Mainzer Stadtwerke AG": 5, "KSB AG": 4, "Zahnen Technik GmbH": 4, "Deutsche Bundesbank, Hauptverwaltung in Rheinland-Pfalz und dem Saarland": 4, "Sparkassenverband Rheinland-Pfalz": 4, "Landesbank Baden-Württemberg (LBBW)": 4, "Debeka Versicherungsgruppe": 4, "Provinzial Versicherung AG": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 5, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Universität Trier": 4, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 5, "Hochschule Mainz": 3, "Empolis Information Management GmbH": 4, "BRICKMAKERS AG": 4},
    "digitale_bau_verwaltung_planung": {"vero – Verband der Bau- und Rohstoffindustrie": 5, "Bauwirtschaft Rheinland-Pfalz e.V.": 5, "Architektenkammer Rheinland-Pfalz": 5, "Ingenieurkammer Rheinland-Pfalz": 5, "Städtetag Rheinland-Pfalz": 5, "Landkreistag Rheinland-Pfalz": 5, "Deutsche Universität für Verwaltungswissenschaften Speyer": 4, "Hochschule Mainz": 5, "Technische Hochschule Bingen": 4, "Hochschule Kaiserslautern": 3, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 4, "Empolis Information Management GmbH": 4, "Aturas GmbH": 4, "BRICKMAKERS AG": 4, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 4, "Karl Gemünden GmbH & Co. KG": 3, "Heberger Gruppe": 3, "Entwicklungsagentur Rheinland-Pfalz e.V.": 3},
    "kultur_standort_identitaet": {"Villa Musica Rheinland-Pfalz": 5, "SWR – Südwestrundfunk": 5, "ZDF – Zweites Deutsches Fernsehen": 5, "RPR1.": 4, "Städtetag Rheinland-Pfalz": 5, "Landkreistag Rheinland-Pfalz": 4, "Entwicklungsagentur Rheinland-Pfalz e.V.": 5, "1. FSV Mainz 05 e.V.": 4, "Erster FC Kaiserslautern GmbH & Co. KGaA": 4, "Johannes Gutenberg-Universität Mainz": 3, "Hochschule Mainz": 3, "Universität Trier": 3, "Evangelische Kirche der Pfalz": 4, "Caritasverband für die Diözese Speyer e.V.": 3, "LOTTO Rheinland-Pfalz GmbH": 3, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 3, "Staatskanzlei des Landes Rheinland-Pfalz": 4},
    "kreativwirtschaft_digitale_medien": {"SWR – Südwestrundfunk": 5, "ZDF – Zweites Deutsches Fernsehen": 5, "RPR1.": 5, "Hochschule Mainz": 5, "Johannes Gutenberg-Universität Mainz": 4, "Universität Trier": 3, "robotspaceship": 4, "BRICKMAKERS AG": 4, "Deutsche Telekom AG": 4, "Villa Musica Rheinland-Pfalz": 4, "WHU – Otto Beisheim School of Management": 3, "Staatskanzlei des Landes Rheinland-Pfalz": 4, "Entwicklungsagentur Rheinland-Pfalz e.V.": 3, "1. FSV Mainz 05 e.V.": 3, "Erster FC Kaiserslautern GmbH & Co. KGaA": 3},
    "musikstipendium_talentfoerderung": {"Villa Musica Rheinland-Pfalz": 5, "SWR – Südwestrundfunk": 4, "ZDF – Zweites Deutsches Fernsehen": 4, "RPR1.": 3, "Johannes Gutenberg-Universität Mainz": 4, "Hochschule Mainz": 4, "Universität Trier": 3, "Katholische Hochschule Mainz": 3, "Vinzenz Pallotti University gGmbH": 3, "Staatskanzlei des Landes Rheinland-Pfalz": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 3, "LOTTO Rheinland-Pfalz GmbH": 3, "Evangelische Kirche der Pfalz": 3, "Caritasverband für die Diözese Speyer e.V.": 3},
    "hochschule_wirtschaft_transfer": {"Johannes Gutenberg-Universität Mainz": 5, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 5, "WHU – Otto Beisheim School of Management": 5, "SmartFactory KL e.V.": 5, "Universität Trier": 4, "Universität Koblenz": 4, "Hochschule Mainz": 4, "Hochschule Trier": 4, "Hochschule Koblenz": 4, "Hochschule Kaiserslautern": 4, "Hochschule Worms": 4, "Hochschule für Wirtschaft und Gesellschaft Ludwigshafen": 4, "Technische Hochschule Bingen": 4, "Duale Hochschule Rheinland-Pfalz": 4, "Katholische Hochschule Mainz": 3, "Vinzenz Pallotti University gGmbH": 3, "BASF SE": 4, "Boehringer Ingelheim Pharma GmbH & Co. KG": 4, "SCHOTT AG": 4, "ITK Engineering GmbH": 4, "Empolis Information Management GmbH": 4, "BRICKMAKERS AG": 4, "Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 3, "Handwerkskammern Rheinland-Pfalz": 4, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 4},
    "weiterbildung_lebenslanges_lernen": {"Duale Hochschule Rheinland-Pfalz": 5, "Bundesagentur für Arbeit, Regionaldirektion Rheinland-Pfalz-Saarland": 5, "Handwerkskammern Rheinland-Pfalz": 5, "Deutscher Gewerkschaftsbund (DGB) Rheinland-Pfalz/Saarland": 4, "Hochschule für Wirtschaft und Gesellschaft Ludwigshafen": 4, "Hochschule Kaiserslautern": 4, "Hochschule Koblenz": 4, "Hochschule Mainz": 4, "Hochschule Trier": 4, "Hochschule Worms": 4, "Technische Hochschule Bingen": 4, "Katholische Hochschule Mainz": 4, "Vinzenz Pallotti University gGmbH": 4, "WHU – Otto Beisheim School of Management": 4, "Alexianer GmbH": 4, "Barmherzige Brüder Trier gGmbH": 4, "Caritasverband für die Diözese Speyer e.V.": 4, "JOBS FOR MOMS UG": 4, "Leonardo PersonalKonzept GmbH": 3, "GLOBUS Markthallen Holding GmbH & Co. KG": 3, "BASF SE": 3, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 3},
    "wissenschaft_demokratie_gesellschaft": {"Johannes Gutenberg-Universität Mainz": 5, "Universität Trier": 5, "Universität Koblenz": 4, "Deutsche Universität für Verwaltungswissenschaften Speyer": 5, "Hochschule Mainz": 4, "Katholische Hochschule Mainz": 4, "Vinzenz Pallotti University gGmbH": 4, "SWR – Südwestrundfunk": 5, "ZDF – Zweites Deutsches Fernsehen": 5, "RPR1.": 4, "Evangelische Kirche der Pfalz": 5, "Caritasverband für die Diözese Speyer e.V.": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 5, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Villa Musica Rheinland-Pfalz": 3, "Entwicklungsagentur Rheinland-Pfalz e.V.": 3},
    "verwaltungswissen_digitalisierung": {"Deutsche Universität für Verwaltungswissenschaften Speyer": 5, "Staatskanzlei des Landes Rheinland-Pfalz": 5, "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau": 5, "Städtetag Rheinland-Pfalz": 5, "Landkreistag Rheinland-Pfalz": 5, "Empolis Information Management GmbH": 5, "Aturas GmbH": 4, "BRICKMAKERS AG": 4, "Deutsche Telekom AG": 4, "Universität Trier": 4, "Universität Koblenz": 4, "Johannes Gutenberg-Universität Mainz": 4, "Hochschule Mainz": 4, "Rheinland-Pfälzische Technische Universität Kaiserslautern-Landau (RPTU)": 4, "TÜV Rheinland Berlin Brandenburg Pfalz e.V.": 4, "Investitions- und Strukturbank Rheinland-Pfalz (ISB)": 3, "Sparkassenverband Rheinland-Pfalz": 3},
    "gesundheit_versorgung_praevention": {"AOK Rheinland-Pfalz/Saarland": 5, "IKK Südwest": 5, "Techniker Krankenkasse": 5, "Kassenärztliche Vereinigung Rheinland-Pfalz": 5, "Kassenzahnärztliche Vereinigung Rheinland-Pfalz": 4, "Landesärztekammer Rheinland-Pfalz": 5, "LandesPsychotherapeutenKammer Rheinland-Pfalz": 4, "Alexianer GmbH": 4, "Barmherzige Brüder Trier gGmbH": 4, "Caritasverband für die Diözese Speyer e.V.": 4, "Boehringer Ingelheim Pharma GmbH & Co. KG": 4, "Katholische Hochschule Mainz": 3, "Vinzenz Pallotti University gGmbH": 3, "Johannes Gutenberg-Universität Mainz": 3, "Universität Trier": 3, "Hochschule für Wirtschaft und Gesellschaft Ludwigshafen": 3, "SWR – Südwestrundfunk": 3, "ZDF – Zweites Deutsches Fernsehen": 3, "Staatskanzlei des Landes Rheinland-Pfalz": 3, "Städtetag Rheinland-Pfalz": 3, "Landkreistag Rheinland-Pfalz": 3},
    "pflege_fachkraefte_sozialwirtschaft": {"Caritasverband für die Diözese Speyer e.V.": 5, "Alexianer GmbH": 5, "Barmherzige Brüder Trier gGmbH": 5, "in.betrieb gGmbH": 5, "Katholische Hochschule Mainz": 5, "Vinzenz Pallotti University gGmbH": 4, "Bundesagentur für Arbeit, Regionaldirektion Rheinland-Pfalz-Saarland": 5, "Duale Hochschule Rheinland-Pfalz": 4, "AOK Rheinland-Pfalz/Saarland": 4, "IKK Südwest": 4, "Techniker Krankenkasse": 4, "Kassenärztliche Vereinigung Rheinland-Pfalz": 4, "Landesärztekammer Rheinland-Pfalz": 4, "LandesPsychotherapeutenKammer Rheinland-Pfalz": 4, "Deutscher Gewerkschaftsbund (DGB) Rheinland-Pfalz/Saarland": 4, "JOBS FOR MOMS UG": 3, "Leonardo PersonalKonzept GmbH": 3, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 3},
    "teilhabe_inklusion_arbeitsmarkt": {"JOBS FOR MOMS UG": 5, "in.betrieb gGmbH": 5, "Bundesagentur für Arbeit, Regionaldirektion Rheinland-Pfalz-Saarland": 5, "Deutscher Gewerkschaftsbund (DGB) Rheinland-Pfalz/Saarland": 4, "Caritasverband für die Diözese Speyer e.V.": 5, "Evangelische Kirche der Pfalz": 4, "Alexianer GmbH": 4, "Barmherzige Brüder Trier gGmbH": 4, "Katholische Hochschule Mainz": 4, "Vinzenz Pallotti University gGmbH": 4, "Leonardo PersonalKonzept GmbH": 4, "GLOBUS Markthallen Holding GmbH & Co. KG": 3, "Handwerkskammern Rheinland-Pfalz": 3, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Staatskanzlei des Landes Rheinland-Pfalz": 3, "SWR – Südwestrundfunk": 3, "ZDF – Zweites Deutsches Fernsehen": 3},
    "demokratie_medien_zusammenhalt": {"SWR – Südwestrundfunk": 5, "ZDF – Zweites Deutsches Fernsehen": 5, "RPR1.": 5, "Evangelische Kirche der Pfalz": 5, "Caritasverband für die Diözese Speyer e.V.": 4, "Johannes Gutenberg-Universität Mainz": 4, "Universität Trier": 4, "Universität Koblenz": 4, "Deutsche Universität für Verwaltungswissenschaften Speyer": 5, "Hochschule Mainz": 3, "Katholische Hochschule Mainz": 4, "Vinzenz Pallotti University gGmbH": 3, "Staatskanzlei des Landes Rheinland-Pfalz": 5, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Villa Musica Rheinland-Pfalz": 3, "Entwicklungsagentur Rheinland-Pfalz e.V.": 3, "1. FSV Mainz 05 e.V.": 3, "Erster FC Kaiserslautern GmbH & Co. KGaA": 3},
    "sport_region_identifikation": {"1. FSV Mainz 05 e.V.": 5, "Erster FC Kaiserslautern GmbH & Co. KGaA": 5, "SWR – Südwestrundfunk": 4, "ZDF – Zweites Deutsches Fernsehen": 4, "RPR1.": 4, "AOK Rheinland-Pfalz/Saarland": 4, "IKK Südwest": 4, "Techniker Krankenkasse": 4, "Städtetag Rheinland-Pfalz": 4, "Landkreistag Rheinland-Pfalz": 4, "Johannes Gutenberg-Universität Mainz": 3, "Hochschule Mainz": 3, "Hochschule Kaiserslautern": 3, "Universität Trier": 3, "LOTTO Rheinland-Pfalz GmbH": 4, "GLOBUS Markthallen Holding GmbH & Co. KG": 3, "Evangelische Kirche der Pfalz": 3, "Caritasverband für die Diözese Speyer e.V.": 3},
}

PUBLIC_SYSTEM_MEMBERS = {
    "Bundesagentur für Arbeit, Regionaldirektion Rheinland-Pfalz-Saarland",
    "Deutsche Bundesbank, Hauptverwaltung in Rheinland-Pfalz und dem Saarland",
    "Deutsche Universität für Verwaltungswissenschaften Speyer",
    "Duale Hochschule Rheinland-Pfalz",
    "Handwerkskammern Rheinland-Pfalz",
    "Investitions- und Strukturbank Rheinland-Pfalz (ISB)",
    "Landesärztekammer Rheinland-Pfalz",
    "LandesPsychotherapeutenKammer Rheinland-Pfalz",
    "Landkreistag Rheinland-Pfalz",
    "Ministerium für Wirtschaft, Verkehr, Landwirtschaft und Weinbau",
    "Staatskanzlei des Landes Rheinland-Pfalz",
    "Städtetag Rheinland-Pfalz",
    "Steuerberaterkammer Rheinland-Pfalz",
}

BROAD_MEDIA_MEMBERS_AUDIT = {
    "SWR – Südwestrundfunk",
    "SWR - Südwestrundfunk",
    "ZDF – Zweites Deutsches Fernsehen",
    "ZDF - Zweites Deutsches Fernsehen",
    "RPR1.",
}


def build_member_environment_weights() -> dict[str, dict[str, float]]:
    member_scores: dict[str, list[int]] = defaultdict(list)
    for relevance in CLUSTER_MEMBER_RELEVANCE.values():
        for member, score in relevance.items():
            member_scores[member].append(int(score))

    weights = {}
    for member, scores in member_scores.items():
        max_score = max(scores)
        avg_score = sum(scores) / len(scores)
        breadth = min(len(scores), 5)
        is_public = member in PUBLIC_SYSTEM_MEMBERS
        is_media = member in BROAD_MEDIA_MEMBERS_AUDIT
        weights[member] = {
            "economic_weight": max_score,
            "regional_system_role": max(max_score, 5 if is_public else 3 if is_media else round(avg_score)),
            "network_multiplier": min(5, 2 + breadth),
            "innovation_relevance": round(avg_score),
            "public_interest": 5 if is_public or is_media else max(1, min(5, round(avg_score))),
        }
    return weights


MEMBER_ENVIRONMENT_WEIGHTS = build_member_environment_weights()



LOW_VALUE_TITLE_PATTERNS = [
    "ladendiebstahl",
    "so tanken sie",
    "waffenruhe wackelt",
    "ölpreis",
    "bremsspuren des krieges",
    "infoabend",
    "nacht der wissenschaft",
    "girls'day",
    "girlsday",
    "zurück am campus",
    "fischdetektiv",
    "das beständige haus",
]

LOW_VALUE_TITLE_PATTERNS.extend([
    "polizei",
    "verkehrsunfall",
    "tanken sie",
    "usa und iran",
    "ölpreis",
    "oelpreis",
    "allzeithoch",
    "marktkommentar",
    "sentix",
    "zurück am campus",
    "das beständige haus",
    "ticket",
    "podcast",
    "folge",
    "sendung",
    "sponsoring",
    "brustring",
    "vfb stuttgart",
    "automatisiert erstellt",
    "gewinnspiel",
    "dissertation",
    "dissertation erfolgreich verteidigt",
    "ranking",
    "che-ranking",
    "bestnoten",
    "auszeichnung",
])


@dataclass
class MemberProfile:
    member_id: int
    name: str
    homepage: str
    topics: Counter
    sections: Counter
    clusters: Counter
    names: Counter
    titles: list[str]
    urls: list[str]
    signals: list[dict[str, str]]
    event_count: int = 0
    avg_decision_score: float = 0.0
    avg_event_score: float = 0.0


class EvidenceLevel(str, Enum):
    NOISE = "noise"
    CONTEXT = "context"
    CONVENING = "convening"


class ConveningTypology(str, Enum):
    MAP_LANDSCAPE = "Map a Landscape"
    FORGE_ALLIANCES = "Forge Alliances"
    DISCOVER_PATH_FORWARD = "Discover a Path Forward"
    ACCELERATE_ACTION = "Accelerate Action"
    SHARE_LEARNING = "Share Learning"
    INVITE_LEARNING = "Invite Learning"
    AMPLIFY_MESSAGE = "Amplify a Message"
    LEVERAGE_GLOBAL_MOMENT = "Leverage a Global Moment"


class PurposeMaturity(str, Enum):
    SIGNAL = "signal"
    PROBLEM = "problem"
    OPPORTUNITY = "opportunity"
    PATHWAY = "pathway"
    COMMITMENT = "commitment"


class ConveningStage(str, Enum):
    MAP_LANDSCAPE = "map_landscape"
    FORGE_ALLIANCES = "forge_alliances"
    DISCOVER_PATH_FORWARD = "discover_path_forward"
    ACCELERATE_ACTION = "accelerate_action"
    SHARE_LEARNING = "share_learning"
    INVITE_LEARNING = "invite_learning"
    AMPLIFY_MESSAGE = "amplify_message"
    LEVERAGE_GLOBAL_MOMENT = "leverage_global_moment"


class NextEngagementMove(str, Enum):
    INTERNAL_BRIEF = "internal_brief"
    VALIDATE_SIGNAL = "validate_signal"
    BILATERAL_SOUNDING = "bilateral_sounding"
    TRIAD_BRIDGE = "triad_bridge"
    LANDSCAPE_SESSION = "landscape_session"
    PATH_FORWARD_WORKSHOP = "path_forward_workshop"
    COMMITMENT_ROUNDTABLE = "commitment_roundtable"
    PUBLIC_AMPLIFICATION = "public_amplification"


@dataclass(frozen=True)
class ConveningDesignArc:
    north_star_purpose: str
    stage: ConveningStage
    participants_logic: str
    required_inputs: tuple[str, ...]
    expected_outputs: tuple[str, ...]
    participant_takeaway: str
    success_metrics: tuple[str, ...]
    next_decision_gate: str


ACTOR_USE_MODE = {
    "source_only": "appears only as source/context",
    "visibility_actor": "can amplify or frame public debate",
    "implementation_actor": "can act directly",
    "resource_actor": "can fund, staff, host, or authorize",
    "knowledge_actor": "can supply expertise or evidence",
    "problem_owner": "owns or directly feels the problem",
    "relationship_broker": "can connect actors or convene trust",
}


STAGE_OUTPUT_RULES = {
    ConveningStage.MAP_LANDSCAPE: {
        "allowed_outputs": {"problem_map", "opportunity_map", "actor_map"},
        "forbidden_claims": {"commitment", "pilot_launch", "funding_secured"},
    },
    ConveningStage.FORGE_ALLIANCES: {
        "allowed_outputs": {"relationship_map", "shared_interest", "follow_up_group"},
        "forbidden_claims": {"implementation_plan", "secured_resources"},
    },
    ConveningStage.DISCOVER_PATH_FORWARD: {
        "allowed_outputs": {"pilot_question", "decision_options", "roadmap_recommendations"},
        "forbidden_claims": {"final_commitments"},
    },
    ConveningStage.ACCELERATE_ACTION: {
        "allowed_outputs": {"owners_named", "resources_committed", "implementation_plan"},
        "forbidden_claims": set(),
    },
    ConveningStage.SHARE_LEARNING: {
        "allowed_outputs": {"briefing_note", "peer_learning_takeaway"},
        "forbidden_claims": {"implementation_commitment"},
    },
    ConveningStage.INVITE_LEARNING: {
        "allowed_outputs": {"validated_question", "learning_agenda"},
        "forbidden_claims": {"implementation_plan", "funding_secured"},
    },
    ConveningStage.AMPLIFY_MESSAGE: {
        "allowed_outputs": {"message_owner", "visibility_format", "target_audience"},
        "forbidden_claims": {"new_coalition_commitment"},
    },
    ConveningStage.LEVERAGE_GLOBAL_MOMENT: {
        "allowed_outputs": {"agenda_hook", "regional_convening_ask"},
        "forbidden_claims": {"local_implementation_commitment"},
    },
}


TYPOLOGY_REQUIRED_ROLES = {
    ConveningTypology.MAP_LANDSCAPE: {
        "knowledge_actor",
        "implementation_actor",
        "problem_owner",
        "public_interest_voice",
    },
    ConveningTypology.FORGE_ALLIANCES: {
        "relationship_broker",
        "implementation_actor",
        "knowledge_actor",
    },
    ConveningTypology.DISCOVER_PATH_FORWARD: {
        "problem_owner",
        "technical_or_domain_expert",
        "implementation_actor",
        "decision_maker",
    },
    ConveningTypology.ACCELERATE_ACTION: {
        "decision_maker",
        "resource_holder",
        "implementation_actor",
        "coalition_anchor",
    },
    ConveningTypology.SHARE_LEARNING: {
        "content_expert",
        "participant_group",
    },
    ConveningTypology.INVITE_LEARNING: {
        "host_actor",
        "divergent_perspective",
        "knowledge_actor",
    },
    ConveningTypology.AMPLIFY_MESSAGE: {
        "message_owner",
        "media_actor",
        "coalition_anchor",
    },
    ConveningTypology.LEVERAGE_GLOBAL_MOMENT: {
        "agenda_owner",
        "public_authority",
        "media_actor",
        "international_or_policy_context_actor",
    },
}


@dataclass
class MeetingPattern:
    pattern_id: int
    members: tuple[str, ...]
    score: float
    shared_topics: tuple[str, ...]
    bridge_topics: tuple[str, ...]
    concrete_clusters: tuple[str, ...]
    relevant_names: tuple[str, ...]
    reason: str
    convening_theme: str
    why_these_members: str
    editorial_justification: str
    recent_signals: tuple[str, ...]
    suggested_format: str
    possible_agenda: str
    convening_typology: str = ""
    typology_confidence: float = 0.0
    typology_reason: str = ""
    expected_output: str = ""
    required_roles: tuple[str, ...] = ()
    candidate_id: str = ""
    scip_variable: str = ""
    candidate_source: str = "generated"
    opportunity_key: str = ""
    member_pair_key: str = ""
    static_support_score: float = 0.0
    live_signal_boost: float = 0.0
    live_event_count: int = 0
    live_signal_titles: tuple[str, ...] = ()
    live_signal_urls: tuple[str, ...] = ()
    live_evidence_status: str = ""
    geo_proximity_score: float = 0.0
    geo_proximity_norm: float = 0.0
    distance_km: float = 0.0
    same_geo_region: bool = False
    geo_score_reason: str = ""
    candidate_status: str = ""
    filter_reason: str = ""
    weak_evidence_flag: int = 0
    low_score_flag: int = 0
    academic_pair_flag: int = 0
    academic_academic_flag: int = 0
    non_academic_pair_flag: int = 0
    legacy_candidate_id: str = ""
    legacy_opportunity_key: str = ""
    old_candidate_max_score: float = 0.0
    old_candidate_rows: int = 0


def latest_file(prefix: str, directory: Path = REPORT_DIR) -> Path:
    files = sorted(directory.glob(f"{prefix}_*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No files found for prefix {prefix!r} in {directory}")
    return files[0]


def read_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def actor_name_key(member_name: str) -> str:
    return network_member_key(member_name)


def is_academic_support_member(member_name: str) -> bool:
    key = actor_name_key(member_name)
    return any(marker in key for marker in ACADEMIC_MEMBER_MARKERS)


def is_profit_oriented_member(member_name: str) -> bool:
    key = actor_name_key(member_name)
    if is_academic_support_member(member_name):
        return False
    if " e.v." in key or " ev " in key:
        return False
    return any(marker in key for marker in BUSINESS_MEMBER_MARKERS)


def is_implementation_anchor_member(member_name: str) -> bool:
    key = actor_name_key(member_name)
    if is_academic_support_member(member_name):
        return False
    return is_profit_oriented_member(member_name) or any(marker in key for marker in IMPLEMENTATION_MEMBER_MARKERS)


def actor_role(member_name: str) -> str:
    configured_role = NETWORK_MEMBER_ROLES.get(actor_name_key(member_name))
    if configured_role:
        return configured_role
    if is_profit_oriented_member(member_name):
        return "profit_anchor"
    if is_implementation_anchor_member(member_name):
        return "implementation_anchor"
    if is_academic_support_member(member_name):
        return "academic_support"
    return "context_actor"


def actor_has_public_role(member_name: str, evidence_text: str) -> bool:
    role = actor_role(member_name)
    if role in {"profit_anchor", "implementation_anchor", "academic_support"}:
        return True
    actor_lower = clean_display_value(member_name).lower()
    text = clean_display_value(evidence_text).lower()
    if "swr" in actor_lower:
        return any(term in text for term in ["versorgung", "öffentlich", "oeffentlich", "kommunen", "problem", "studie", "alarm", "debatte", "sichtbarkeit"])
    if "zdf" in actor_lower:
        return any(term in text for term in ["öffentlich", "oeffentlich", "debatte", "gesellschaft", "sichtbarkeit", "demokratie"])
    if "mainz 05" in actor_lower or "fsv mainz" in actor_lower:
        return any(term in text for term in ["jugend", "prävention", "praevention", "gesund", "community", "bildung", "spende", "stadion", "fans"])
    return False


def pattern_has_public_actor_roles(row: pd.Series) -> bool:
    members = [part.strip() for part in clean_display_value(row.get("members", "")).split("|") if part.strip()]
    evidence = " ".join(
        clean_display_value(row.get(key, ""))
        for key in [
            "convening_theme",
            "shared_topics",
            "bridge_topics",
            "concrete_clusters",
            "recent_signals",
            "editorial_justification",
        ]
    )
    return bool(members) and all(actor_has_public_role(member, evidence) for member in members)


def actor_role_counts(member_names: tuple[str, ...] | list[str]) -> dict[str, int]:
    counts = Counter(actor_role(name) for name in member_names)
    return {
        "profit_anchor": counts.get("profit_anchor", 0),
        "implementation_anchor": counts.get("implementation_anchor", 0),
        "academic_support": counts.get("academic_support", 0),
        "context_actor": counts.get("context_actor", 0),
    }


def academic_support_penalty(member_names: tuple[str, ...] | list[str], *, final: bool = False) -> float:
    if not member_names:
        return 0.0
    counts = actor_role_counts(member_names)
    academic = counts["academic_support"]
    anchors = counts["profit_anchor"] + counts["implementation_anchor"]
    if academic <= 0:
        return 0.0
    if anchors > 0:
        # Hochschulen become useful support actors once a practice/business anchor is present.
        return 0.0 if not final else max(0.0, academic - anchors) * 1.5
    if academic == len(member_names):
        return 8.0 if not final else 18.0 + max(0, len(member_names) - 2) * 3.0
    return 4.5 if not final else 10.0


def ordered_members_for_display(member_names: list[str] | tuple[str, ...]) -> list[str]:
    role_rank = {
        "profit_anchor": 0,
        "implementation_anchor": 1,
        "context_actor": 2,
        "academic_support": 3,
    }
    return sorted(member_names, key=lambda name: (role_rank.get(actor_role(name), 9), str(name).lower()))


def slug_key(value: str, *, limit: int = 64) -> str:
    text = repair_mojibake(str(value or "")).lower()
    text = (
        text.replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
    )
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return text[:limit].strip("-") or "unknown"


def opportunity_actor_key(member_names: tuple[str, ...] | list[str]) -> str:
    core = [name for name in member_names if actor_role(name) != "context_actor"]
    actors = core or list(member_names)
    parts = [slug_key(name, limit=28) for name in ordered_members_for_display(actors)]
    return "_".join(parts[:4]) or "unknown-actors"


def opportunity_topic_key_from_values(*values: Any) -> str:
    text = " ".join(clean_display_value(value) for value in values if clean_display_value(value))
    words = [
        word for word in re.findall(r"[a-zA-ZÄÖÜäöüß0-9][a-zA-ZÄÖÜäöüß0-9\-]{2,}", text.lower())
        if word not in {"und", "oder", "mit", "der", "die", "das", "von", "fuer", "für", "zur", "zum", "eine", "einer"}
    ]
    if not words:
        return "general"
    return slug_key("-".join(dict.fromkeys(words[:5])), limit=72)


def opportunity_id_for_pattern(pattern: MeetingPattern) -> str:
    # v2: for Excel-defined candidates, the opportunity_key is the auditable
    # grouping key used for duplicate-opportunity constraints.
    if getattr(pattern, "opportunity_key", ""):
        return str(pattern.opportunity_key)
    actor_key = opportunity_actor_key(pattern.members)
    topic_key = opportunity_topic_key_from_values(
        " ".join(display_topic(t) for t in pattern.concrete_clusters),
        " ".join(display_topic(t) for t in pattern.shared_topics[:3]),
        pattern.convening_theme,
    )
    return f"{actor_key}_{topic_key}"


def load_local_env() -> None:
    for env_path in [Path(".env"), PROJECT_DIR / ".env", PROJECT_DIR.parent / ".env", REPORT_DIR / ".env"]:
        if not env_path.exists():
            continue
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


class LocalOllamaChatClient:
    """Tiny OpenAI-like wrapper for Ollama's local /api/chat endpoint."""

    class _Message:
        def __init__(self, content: str):
            self.content = content

    class _Choice:
        def __init__(self, content: str):
            self.message = LocalOllamaChatClient._Message(content)

    class _Response:
        def __init__(self, content: str):
            self.choices = [LocalOllamaChatClient._Choice(content)]

    class _Completions:
        def __init__(self, outer: "LocalOllamaChatClient"):
            self.outer = outer

        def create(
            self,
            *,
            model: str,
            messages: list[dict[str, str]],
            temperature: float = 0.2,
            max_tokens: int = 360,
            extra_body: Optional[dict[str, Any]] = None,
            **_: Any,
        ) -> "LocalOllamaChatClient._Response":
            options = {
                "temperature": temperature,
                "num_predict": max_tokens,
                "top_p": 0.9,
                "repeat_penalty": 1.1,
            }
            if extra_body and isinstance(extra_body.get("options"), dict):
                options.update(extra_body["options"])
            payload = {
                "model": model,
                "messages": messages,
                "stream": False,
                "format": "json",
                "options": options,
            }
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            request = urllib.request.Request(
                self.outer.chat_url,
                data=data,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(request, timeout=self.outer.timeout) as response:
                raw = response.read().decode("utf-8", errors="replace")
            parsed = json.loads(raw)
            content = ""
            if isinstance(parsed, dict):
                message = parsed.get("message") or {}
                if isinstance(message, dict):
                    content = str(message.get("content", "") or "")
                if not content:
                    content = str(parsed.get("response", "") or "")
            return LocalOllamaChatClient._Response(content)

    class _Chat:
        def __init__(self, outer: "LocalOllamaChatClient"):
            self.completions = LocalOllamaChatClient._Completions(outer)

    def __init__(self, base_url: str, timeout: int):
        self.base_url = base_url.rstrip("/")
        self.chat_url = f"{self.base_url}/api/chat"
        self.tags_url = f"{self.base_url}/api/tags"
        self.timeout = timeout
        self.chat = LocalOllamaChatClient._Chat(self)

    def ping(self) -> bool:
        try:
            with urllib.request.urlopen(self.tags_url, timeout=min(self.timeout, 5)) as response:
                return 200 <= int(response.status) < 300
        except Exception:
            return False


def build_openai_compatible_client():
    if not USE_OLLAMA_FOR_CONVENING_TEXT:
        print("KI Convening Text: nicht aktiv (USE_OLLAMA_FOR_CONVENING_TEXT=False)")
        return None
    load_local_env()
    model = (os.getenv("ZIRP_CONVENING_MODEL", "").strip() or os.getenv("OLLAMA_CONVENING_MODEL", "").strip() or "qwen2.5:7b-instruct")
    base_url = os.getenv("OPENAI_BASE_URL", "").strip()
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if OpenAI is not None:
        if base_url and not api_key:
            api_key = "ollama"
        if api_key:
            try:
                endpoint = base_url or "OpenAI API"
                print(f"KI Convening Text: aktiv ({model}, endpoint={endpoint})")
                return OpenAI(api_key=api_key, base_url=base_url or None, timeout=OLLAMA_CONVENING_TIMEOUT)
            except Exception as exc:
                print(f"KI Convening Text: OpenAI-kompatibler Client nicht aktiv ({exc}); versuche Ollama direkt")

    ollama_url = (
        os.getenv("OLLAMA_BASE_URL", "").strip()
        or os.getenv("OLLAMA_HOST", "").strip()
        or "http://127.0.0.1:11434"
    )
    if not ollama_url.startswith(("http://", "https://")):
        ollama_url = f"http://{ollama_url}"
    client = LocalOllamaChatClient(ollama_url, OLLAMA_CONVENING_TIMEOUT)
    if client.ping():
        print(f"KI Convening Text: aktiv ({model}, endpoint={ollama_url}/api/chat, client=ollama-direct)")
        return client
    print(
        "KI Convening Text: nicht aktiv "
        f"(OpenAI-Paket/API nicht verfuegbar und Ollama nicht erreichbar unter {ollama_url})"
    )
    return None


def clean_topic(topic: str) -> str:
    return re.sub(r"\s+", " ", str(topic).strip().lower())


def split_topics(value: Any) -> list[str]:
    if value is None or pd.isna(value):
        return []
    return [clean_topic(x) for x in str(value).split(",") if clean_topic(x)]


def cluster_topic_key(cluster_id: str) -> str:
    return f"cluster:{cluster_id}"


def plain_cluster_id(topic: str) -> str:
    topic = clean_topic(topic)
    return topic.split(":", 1)[1] if topic.startswith("cluster:") else topic


def cluster_info(topic: str) -> dict[str, Any]:
    return CLUSTER_TOPIC_LOOKUP.get(plain_cluster_id(topic), {})


def classify_clusterable_topic(text: str) -> list[dict[str, Any]]:
    haystack = re.sub(r"\s+", " ", str(text or "").lower())
    matches: list[dict[str, Any]] = []
    for field, clusters in ZIRP_CLUSTER_TOPICS.items():
        for cluster_id, cluster in clusters.items():
            hits = [term for term in cluster["signals"] if term.lower() in haystack]
            if len(hits) >= 2:
                matches.append({
                    "field": field,
                    "cluster_id": cluster_id,
                    "topic": cluster_topic_key(cluster_id),
                    "label": cluster["label"],
                    "problem": cluster["problem"],
                    "hits": hits,
                    "strength": len(hits),
                })
    return sorted(matches, key=lambda x: (x["strength"], PRIORITY_CLUSTER_WEIGHTS.get(x["cluster_id"], 6.0)), reverse=True)


def is_clusterable_news_item(text: str) -> bool:
    haystack = str(text or "").lower()
    if any(noise in haystack for noise in CLUSTER_NOISE_TERMS):
        return False
    return bool(classify_clusterable_topic(haystack))


def normalized_hit_count(terms: set[str] | list[str] | tuple[str, ...], haystack: str) -> int:
    return sum(1 for term in terms if normalize_search_text(term) in haystack)


def is_low_value_event(title: str, snippet: str = "") -> bool:
    haystack = normalize_search_text(f"{title} {snippet}")
    if any(normalize_search_text(pattern) in haystack for pattern in LOW_VALUE_TITLE_PATTERNS):
        return True

    noise_hits = normalized_hit_count(MARKET_NOISE_TERMS, haystack)
    pressure_hits = normalized_hit_count(MARKET_PRESSURE_TERMS, haystack)
    action_hits = normalized_hit_count(ACTION_TERMS, haystack)

    if noise_hits >= 2 and pressure_hits == 0 and action_hits == 0:
        return True
    if len(str(title).split()) <= 2 and action_hits == 0 and pressure_hits == 0:
        return True
    return False


def classify_convening_evidence(signal: dict[str, str]) -> EvidenceLevel:
    text = normalize_search_text(f"{signal.get('title', '')} {signal.get('snippet', '')}")

    macro_hits = normalized_hit_count(MACRO_TERMS, text)
    action_hits = normalized_hit_count(ACTION_TERMS, text)
    rlp_hits = normalized_hit_count(["rheinland-pfalz", "rlp", "mainz", "trier", "koblenz", "pfalz"], text)
    problem_hits = normalized_hit_count(PROBLEM_TERMS, text)
    pressure_hits = normalized_hit_count(MARKET_PRESSURE_TERMS, text)

    if macro_hits and not action_hits and not rlp_hits and not pressure_hits:
        return EvidenceLevel.NOISE
    if (rlp_hits or pressure_hits) and action_hits and problem_hits:
        return EvidenceLevel.CONVENING
    if action_hits or problem_hits or pressure_hits:
        return EvidenceLevel.CONTEXT
    return EvidenceLevel.NOISE


def is_convening_evidence(signal: dict[str, str]) -> bool:
    return classify_convening_evidence(signal) == EvidenceLevel.CONVENING


def signal_evidence_text(value: Any) -> str:
    if isinstance(value, dict):
        keys = [
            "title",
            "snippet",
            "mechanism",
            "named_actors",
            "concrete_assets",
            "next_date_or_window",
            "implementation_angle",
            "scip_reasoning",
            "opportunity_strength",
            "recommended_zirp_use",
            "briefing_sentence",
            "opportunity_keywords",
            "section",
            "topics",
            "clusters",
            "cluster_labels",
            "member",
        ]
        return " ".join(str(value.get(key, "") or "") for key in keys)
    if hasattr(value, "get"):
        keys = [
            "titel",
            "title",
            "snippet",
            "mechanism",
            "named_actors",
            "concrete_assets",
            "next_date_or_window",
            "implementation_angle",
            "scip_reasoning",
            "opportunity_strength",
            "recommended_zirp_use",
            "briefing_sentence",
            "opportunity_keywords",
            "hauptsektion",
            "themenfelder",
            "mitglied",
        ]
        return " ".join(str(value.get(key, "") or "") for key in keys)
    return str(value or "")


def is_full_convening_evidence(value: Any) -> bool:
    return classify_convening_evidence({"title": signal_evidence_text(value), "snippet": ""}) == EvidenceLevel.CONVENING


def evidence_level_for_value(value: Any) -> EvidenceLevel:
    return classify_convening_evidence({"title": signal_evidence_text(value), "snippet": ""})


def evidence_weight_for_level(level: EvidenceLevel) -> float:
    if level == EvidenceLevel.CONVENING:
        return 1.0
    if level == EvidenceLevel.CONTEXT:
        return 0.55
    return 0.0


def is_profile_usable_evidence(value: Any) -> bool:
    return evidence_weight_for_level(evidence_level_for_value(value)) > 0


def has_clear_convening_problem_text(text: str) -> bool:
    normalized = normalize_search_text(text)

    vague_hits = normalized_hit_count(VAGUE_PHRASES, normalized)
    problem_meta = weighted_problem_hit_score(normalized)
    action_meta = weighted_action_hit_score(normalized, verbs=True)
    concrete_hits = problem_meta["score"]
    action_hits = action_meta["score"]
    static_problem_hits = problem_meta["static"] + problem_meta["feedback"]

    if vague_hits >= 2 and concrete_hits < 3:
        return False

    return concrete_hits >= 2 and action_hits >= 1 and static_problem_hits >= 1


def has_clear_convening_problem(pattern: MeetingPattern) -> bool:
    text = " ".join([
        str(pattern.convening_theme),
        str(pattern.why_these_members),
        str(pattern.editorial_justification),
        str(pattern.possible_agenda),
        " ".join(pattern.recent_signals),
    ])
    return has_clear_convening_problem_text(text)


def has_existing_coalition_signal(pattern: MeetingPattern) -> bool:
    text = normalize_search_text(" ".join([
        pattern.reason,
        pattern.editorial_justification,
        pattern.possible_agenda,
        " ".join(pattern.recent_signals),
    ]))
    return any(term in text for term in ["kooperation", "partnerschaft", "allianz", "netzwerk", "arbeitsgruppe", "projektstart", "startet"])


def has_resource_signal(pattern: MeetingPattern) -> bool:
    text = normalize_search_text(" ".join([
        pattern.reason,
        pattern.editorial_justification,
        pattern.possible_agenda,
        " ".join(pattern.recent_signals),
    ]))
    return any(term in text for term in ["forderung", "foerderung", "finanzierung", "budget", "investition", "ressource", "personal", "entscheidung"])


def has_launch_or_publication_signal(pattern: MeetingPattern) -> bool:
    text = normalize_search_text(" ".join(pattern.recent_signals))
    return any(term in text for term in ["publikation", "bericht", "studie", "launch", "kampagne", "presse", "veroffentlicht", "veroeffentlicht"])


def infer_purpose_maturity(
    evidence_count: int,
    distinct_evidence_members: int,
    common_problem_quality: str,
    action_readiness: str,
    has_existing_coalition: bool,
    has_resource_signal_value: bool,
) -> PurposeMaturity:
    readiness = normalize_search_text(action_readiness)
    quality = normalize_search_text(common_problem_quality)
    if has_existing_coalition and has_resource_signal_value and readiness == "hoch":
        return PurposeMaturity.COMMITMENT
    if quality == "stark" and readiness in {"hoch", "mittel"}:
        return PurposeMaturity.PATHWAY
    if distinct_evidence_members >= 2 and evidence_count >= 2:
        return PurposeMaturity.OPPORTUNITY
    if quality in {"stark", "mittel"} or evidence_count >= 2:
        return PurposeMaturity.PROBLEM
    return PurposeMaturity.SIGNAL


def infer_convening_typology(
    evidence_count: int,
    distinct_evidence_members: int,
    common_problem_quality: str,
    action_readiness: str,
    *,
    has_existing_coalition: bool = False,
    has_resource_signal_value: bool = False,
    has_launch_or_publication_signal_value: bool = False,
) -> ConveningTypology:
    if has_launch_or_publication_signal_value:
        return ConveningTypology.AMPLIFY_MESSAGE
    if has_existing_coalition and has_resource_signal_value:
        return ConveningTypology.ACCELERATE_ACTION
    if normalize_search_text(common_problem_quality) == "stark" and normalize_search_text(action_readiness) in {"hoch", "mittel"}:
        return ConveningTypology.DISCOVER_PATH_FORWARD
    if distinct_evidence_members >= 3 and evidence_count >= 4:
        return ConveningTypology.MAP_LANDSCAPE
    if distinct_evidence_members >= 2:
        return ConveningTypology.FORGE_ALLIANCES
    return ConveningTypology.INVITE_LEARNING


def stage_for_typology(typology: ConveningTypology) -> ConveningStage:
    mapping = {
        ConveningTypology.MAP_LANDSCAPE: ConveningStage.MAP_LANDSCAPE,
        ConveningTypology.FORGE_ALLIANCES: ConveningStage.FORGE_ALLIANCES,
        ConveningTypology.DISCOVER_PATH_FORWARD: ConveningStage.DISCOVER_PATH_FORWARD,
        ConveningTypology.ACCELERATE_ACTION: ConveningStage.ACCELERATE_ACTION,
        ConveningTypology.SHARE_LEARNING: ConveningStage.SHARE_LEARNING,
        ConveningTypology.INVITE_LEARNING: ConveningStage.INVITE_LEARNING,
        ConveningTypology.AMPLIFY_MESSAGE: ConveningStage.AMPLIFY_MESSAGE,
        ConveningTypology.LEVERAGE_GLOBAL_MOMENT: ConveningStage.LEVERAGE_GLOBAL_MOMENT,
    }
    return mapping.get(typology, ConveningStage.INVITE_LEARNING)


def typology_stage_gate(
    typology: ConveningTypology,
    *,
    evidence_count: int,
    distinct_evidence_members: int,
    has_clear_problem: bool,
    member_roles: set[str],
    has_resource_signal_value: bool,
    has_launch_or_publication_signal_value: bool,
) -> tuple[bool, str]:
    has_decision_actor = bool(member_roles & {"decision_maker", "problem_owner"})
    has_resource_actor = bool(member_roles & {"resource_holder", "resource_actor"})

    if typology == ConveningTypology.ACCELERATE_ACTION:
        missing = []
        if evidence_count < 3:
            missing.append("evidence_count>=3")
        if distinct_evidence_members < 2:
            missing.append("distinct_evidence_members>=2")
        if not has_clear_problem:
            missing.append("clear_problem")
        if not has_decision_actor:
            missing.append("decision_actor")
        if not (has_resource_actor or has_resource_signal_value):
            missing.append("resource_actor_or_signal")
        return (not missing, ", ".join(missing))

    if typology == ConveningTypology.DISCOVER_PATH_FORWARD:
        missing = []
        if evidence_count < 2:
            missing.append("evidence_count>=2")
        if distinct_evidence_members < 2:
            missing.append("distinct_evidence_members>=2")
        if not has_clear_problem:
            missing.append("clear_problem")
        if not has_decision_actor:
            missing.append("decision_actor")
        return (not missing, ", ".join(missing))

    if typology == ConveningTypology.MAP_LANDSCAPE:
        missing = []
        if evidence_count < 3:
            missing.append("evidence_count>=3")
        if distinct_evidence_members < 3:
            missing.append("distinct_evidence_members>=3")
        return (not missing, ", ".join(missing))

    if typology == ConveningTypology.FORGE_ALLIANCES:
        ok = distinct_evidence_members >= 2
        return ok, "" if ok else "distinct_evidence_members>=2"

    if typology == ConveningTypology.AMPLIFY_MESSAGE:
        ok = has_launch_or_publication_signal_value
        return ok, "" if ok else "launch_or_publication_signal"

    return True, ""


def apply_typology_stage_gate(
    typology: ConveningTypology,
    *,
    evidence_count: int,
    distinct_evidence_members: int,
    common_problem_quality: str,
    action_readiness: str,
    has_clear_problem: bool,
    member_roles: set[str],
    has_resource_signal_value: bool,
    has_launch_or_publication_signal_value: bool,
) -> tuple[ConveningTypology, str]:
    ok, reason = typology_stage_gate(
        typology,
        evidence_count=evidence_count,
        distinct_evidence_members=distinct_evidence_members,
        has_clear_problem=has_clear_problem,
        member_roles=member_roles,
        has_resource_signal_value=has_resource_signal_value,
        has_launch_or_publication_signal_value=has_launch_or_publication_signal_value,
    )
    if ok:
        return typology, ""

    quality = normalize_search_text(common_problem_quality)
    readiness = normalize_search_text(action_readiness)
    if typology == ConveningTypology.ACCELERATE_ACTION and quality == "stark" and readiness in {"hoch", "mittel"} and has_clear_problem:
        downgraded = ConveningTypology.DISCOVER_PATH_FORWARD
    elif typology not in {ConveningTypology.MAP_LANDSCAPE, ConveningTypology.DISCOVER_PATH_FORWARD} and distinct_evidence_members >= 3 and evidence_count >= 3:
        downgraded = ConveningTypology.MAP_LANDSCAPE
    elif typology != ConveningTypology.FORGE_ALLIANCES and distinct_evidence_members >= 2:
        downgraded = ConveningTypology.FORGE_ALLIANCES
    else:
        downgraded = ConveningTypology.INVITE_LEARNING
    return downgraded, f"{typology.value} downgraded to {downgraded.value}; missing {reason}"


def typology_expected_output(typology: ConveningTypology) -> str:
    outputs = {
        ConveningTypology.MAP_LANDSCAPE: "shared problem map and shortlist of follow-up questions",
        ConveningTypology.FORGE_ALLIANCES: "new working relationship and scoped collaboration question",
        ConveningTypology.DISCOVER_PATH_FORWARD: "decision on next step, owner, and pilot question",
        ConveningTypology.ACCELERATE_ACTION: "named commitment path, resources, and 30-day follow-up",
        ConveningTypology.SHARE_LEARNING: "briefing note or peer-learning takeaway",
        ConveningTypology.INVITE_LEARNING: "validated question and learning agenda",
        ConveningTypology.AMPLIFY_MESSAGE: "clear message owner, visibility format, and audience",
        ConveningTypology.LEVERAGE_GLOBAL_MOMENT: "timely policy or agenda hook with concrete convening ask",
    }
    return outputs.get(typology, "validated next convening step")


def typology_reason_for_pattern(
    typology: ConveningTypology,
    evidence_count: int,
    distinct_evidence_members: int,
    common_problem_quality: str,
    action_readiness: str,
) -> str:
    return (
        f"{typology.value} because evidence_count={evidence_count}, "
        f"distinct_evidence_members={distinct_evidence_members}, "
        f"problem_quality={common_problem_quality}, action_readiness={action_readiness}."
    )


def typology_confidence(
    typology: ConveningTypology,
    evidence_count: int,
    distinct_evidence_members: int,
    has_clear_problem: bool,
) -> float:
    score = 0.35
    score += min(evidence_count, 5) * 0.07
    score += min(distinct_evidence_members, 4) * 0.07
    if has_clear_problem:
        score += 0.15
    if typology == ConveningTypology.ACCELERATE_ACTION:
        score -= 0.08
    return round(max(0.1, min(score, 0.95)), 2)


def recommended_action_for_typology(
    typology: ConveningTypology,
    evidence_quality: str,
    missing_roles: set[str],
) -> str:
    if missing_roles:
        return "brief_or_validate"
    if typology == ConveningTypology.MAP_LANDSCAPE:
        return "curate_landscape_session"
    if typology == ConveningTypology.FORGE_ALLIANCES:
        return "host_sounding_conversation"
    if typology == ConveningTypology.DISCOVER_PATH_FORWARD:
        return "convene_decision_workshop"
    if typology == ConveningTypology.ACCELERATE_ACTION:
        return "secure_commitment_roundtable"
    if typology == ConveningTypology.AMPLIFY_MESSAGE:
        return "prepare_public_launch_or_visibility_format"
    return "watch_or_reframe"


def typology_do_not_ask_for(typology: ConveningTypology) -> str:
    values = {
        ConveningTypology.MAP_LANDSCAPE: "binding commitments before the problem is mapped",
        ConveningTypology.FORGE_ALLIANCES: "a full project plan before trust and complementarity are tested",
        ConveningTypology.DISCOVER_PATH_FORWARD: "broad ideation without a decision question",
        ConveningTypology.ACCELERATE_ACTION: "another exploratory discussion without commitments",
        ConveningTypology.AMPLIFY_MESSAGE: "visibility without a clear message owner",
        ConveningTypology.INVITE_LEARNING: "implementation commitments from a weak signal",
    }
    return values.get(typology, "a broader mandate than the evidence supports")


def next_engagement_move_for_stage(
    stage: ConveningStage,
    *,
    missing_roles: set[str],
    has_clear_problem: bool,
    member_count: int,
) -> NextEngagementMove:
    if missing_roles and not has_clear_problem:
        return NextEngagementMove.VALIDATE_SIGNAL
    if stage == ConveningStage.MAP_LANDSCAPE:
        return NextEngagementMove.LANDSCAPE_SESSION
    if stage == ConveningStage.FORGE_ALLIANCES:
        return NextEngagementMove.TRIAD_BRIDGE if member_count >= 3 else NextEngagementMove.BILATERAL_SOUNDING
    if stage == ConveningStage.DISCOVER_PATH_FORWARD:
        return NextEngagementMove.PATH_FORWARD_WORKSHOP
    if stage == ConveningStage.ACCELERATE_ACTION:
        return NextEngagementMove.COMMITMENT_ROUNDTABLE
    if stage == ConveningStage.AMPLIFY_MESSAGE:
        return NextEngagementMove.PUBLIC_AMPLIFICATION
    if stage == ConveningStage.SHARE_LEARNING:
        return NextEngagementMove.INTERNAL_BRIEF
    return NextEngagementMove.VALIDATE_SIGNAL


def design_arc_for_pattern(
    pattern: MeetingPattern,
    typology: ConveningTypology,
    stage: ConveningStage,
    next_move: NextEngagementMove,
    required_roles: set[str],
    missing_roles: set[str],
) -> ConveningDesignArc:
    theme = clean_display_value(pattern.convening_theme) or "regional opportunity signal"
    member_line = " + ".join(ordered_members_for_display(pattern.members))
    rules = STAGE_OUTPUT_RULES.get(stage, {"allowed_outputs": set(), "forbidden_claims": set()})
    required_inputs = tuple(sorted(required_roles)) or ("validated signal",)
    if missing_roles:
        required_inputs = required_inputs + tuple(f"validate missing role: {role}" for role in sorted(missing_roles))
    expected_outputs = tuple(sorted(rules.get("allowed_outputs", set()))) or (typology_expected_output(typology),)

    return ConveningDesignArc(
        north_star_purpose=f"{typology.value}: {theme} in den nächsten belegbaren Convening-Schritt übersetzen.",
        stage=stage,
        participants_logic=f"{member_line} werden wegen komplementärer Rollen einbezogen, nicht wegen ähnlicher Schlagworte.",
        required_inputs=required_inputs,
        expected_outputs=expected_outputs,
        participant_takeaway=f"Die Teilnehmenden klären, ob {next_move.value} jetzt tragfähig ist.",
        success_metrics=(
            "klarer nächster Owner oder Validierungsfrage",
            "Evidenzlücke ausdrücklich benannt",
            "30-Tage-Folgeentscheidung festgelegt",
        ),
        next_decision_gate="nach der nächsten Validierung: vertiefen, reframen oder beobachten",
    )


def profile_core_topics(profile: MemberProfile) -> set[str]:
    return set(profile.topics) & CORE_CONVENING_TOPICS


def profile_strength(profile: MemberProfile) -> float:
    core_bonus = sum(topic_weight(topic) for topic in profile_core_topics(profile))
    return profile.event_count + profile.avg_decision_score / 4 + core_bonus / 8


def is_broad_media_profile(profile: MemberProfile) -> bool:
    return normalize_search_text(profile.name) in BROAD_MEDIA_MEMBERS_NORM


ORG_MARKERS = {
    "gmbh", "ag", "se", "kg", "hochschule", "universitat", "universitaet",
    "ministerium", "stadt", "kreis", "verband", "kammer", "bank", "sparkasse",
}
PERSON_TITLE_MARKERS = {"prof.", "dr.", "ministerin", "minister", "prasidentin", "praesident", "präsidentin", "präsident"}


def likely_person_name(name: str, context: str = "") -> bool:
    normalized_name = normalize_search_text(name)
    normalized_context = normalize_search_text(context)
    if any(marker in normalized_name for marker in ORG_MARKERS):
        return False
    if any(title in normalized_context for title in PERSON_TITLE_MARKERS):
        return True
    parts = str(name or "").split()
    return 2 <= len(parts) <= 4 and all(len(part) > 2 for part in parts)


def extract_names(text: str) -> list[str]:
    source_text = re.sub(r"\s+", " ", str(text))
    patterns = [
        r"\b(?:Prof\.|Dr\.|Prof\.\s*Dr\.|Prof\.\s*Dr\.-Ing\.)\s+[A-ZÄÖÜ][A-Za-zÄÖÜäöüß\-]+(?:\s+[A-ZÄÖÜ][A-Za-zÄÖÜäöüß\-]+){0,3}",
        r"\b[A-ZÄÖÜ][a-zäöüß\-]{2,}\s+[A-ZÄÖÜ][a-zäöüß\-]{2,}\b",
    ]
    names: list[str] = []
    blacklist_norm = {
        normalize_search_text(name)
        for name in {
            "Rheinland Pfalz", "Otto Beisheim", "Bachelor Science", "Fritz Walter",
            "FC Kaiserslautern", "Johannes Gutenberg", "Zweites Deutsches",
        }
    }
    for pattern in patterns:
        for match in re.findall(pattern, source_text):
            name = re.sub(r"\s+", " ", match).strip(" .,;:")
            if normalize_search_text(name) not in blacklist_norm and likely_person_name(name, source_text):
                names.append(name)
    return names


def filter_events_to_period(events_df: pd.DataFrame, lookback_days: int = LOOKBACK_DAYS) -> pd.DataFrame:
    if events_df.empty or "datum" not in events_df.columns:
        return events_df
    df = events_df.copy()
    df["datum_obj"] = pd.to_datetime(df["datum"], errors="coerce")
    df = df[df["datum_obj"].notna()]
    if df.empty:
        return df
    latest = df["datum_obj"].max().normalize()
    cutoff = latest - timedelta(days=lookback_days - 1)
    df = df[(df["datum_obj"] >= cutoff) & (df["datum_obj"] <= latest)].copy()
    # Prioritize the most recent 7 days while keeping days 8-14 as context.
    age_days = (latest - df["datum_obj"].dt.normalize()).dt.days
    df["recent_priority_weight"] = age_days.apply(
        lambda days: RECENT_SIGNAL_WEIGHT if days < RECENT_PRIORITY_DAYS else 1.0
    )
    return df


def build_member_profiles(members_df: pd.DataFrame, events_df: pd.DataFrame) -> dict[str, MemberProfile]:
    profiles: dict[str, MemberProfile] = {}
    for _, row in members_df.iterrows():
        name = str(row.get("mitglied", "")).strip()
        if not name:
            continue
        profiles[name] = MemberProfile(
            member_id=int(row.get("member_id", 0) or 0),
            name=name,
            homepage=str(row.get("homepage", "") or ""),
            topics=Counter(),
            sections=Counter(),
            clusters=Counter(),
            names=Counter(),
            titles=[],
            urls=[],
            signals=[],
        )

    for _, row in events_df.iterrows():
        member = str(row.get("mitglied", "")).strip()
        if not member:
            continue
        if member not in profiles:
            profiles[member] = MemberProfile(
                member_id=int(row.get("member_id", 0) or 0),
                name=member,
                homepage="",
                topics=Counter(),
                sections=Counter(),
                names=Counter(),
                titles=[],
                urls=[],
                signals=[],
            )

        profile = profiles[member]
        title = str(row.get("titel", "") or "")
        snippet = str(row.get("snippet", "") or "")
        mechanism = str(row.get("mechanism", "") or "")
        named_actors = str(row.get("named_actors", "") or "")
        concrete_assets = str(row.get("concrete_assets", "") or "")
        next_window = str(row.get("next_date_or_window", "") or "")
        implementation_angle = str(row.get("implementation_angle", "") or "")
        scip_reasoning = str(row.get("scip_reasoning", "") or "")
        opportunity_strength = str(row.get("opportunity_strength", "") or "")
        recommended_zirp_use = str(row.get("recommended_zirp_use", "") or "")
        briefing_sentence = str(row.get("briefing_sentence", "") or "")
        opportunity_keywords = str(row.get("opportunity_keywords", "") or "")
        section = clean_topic(row.get("hauptsektion", ""))
        broad_topics = split_topics(row.get("themenfelder", ""))
        cluster_text = " ".join([
            title,
            snippet,
            mechanism,
            named_actors,
            concrete_assets,
            implementation_angle,
            scip_reasoning,
            opportunity_strength,
            recommended_zirp_use,
            briefing_sentence,
            opportunity_keywords,
            section,
            " ".join(broad_topics),
            member,
        ])
        cluster_matches = classify_clusterable_topic(cluster_text)

        if is_low_value_event(title, snippet):
            continue
        evidence_level = evidence_level_for_value(row)
        evidence_weight = evidence_weight_for_level(evidence_level)
        if evidence_weight <= 0:
            continue
        if not cluster_matches:
            continue

        profile.event_count += evidence_weight
        if section:
            profile.sections[section] += 1
            profile.topics[section] += 1
        for topic in broad_topics:
            profile.topics[topic] += 1
        for match in cluster_matches[:3]:
            cluster_key = match["topic"]
            profile.clusters[cluster_key] += match["strength"] * evidence_weight
            profile.topics[cluster_key] += (4 + match["strength"]) * evidence_weight
            profile.topics[match["field"]] += evidence_weight
        for name in extract_names(f"{title} {snippet}"):
            profile.names[name] += evidence_weight
        if title:
            profile.titles.append(title)
        url = str(row.get("url", "") or "")
        if url:
            profile.urls.append(url)
        profile.signals.append({
            "date": str(row.get("datum", "") or ""),
            "title": title,
            "section": section,
            "topics": ", ".join(broad_topics),
            "clusters": ", ".join(match["topic"] for match in cluster_matches[:3]),
            "cluster_labels": ", ".join(match["label"] for match in cluster_matches[:3]),
            "snippet": re.sub(r"\s+", " ", snippet).strip()[:220],
            "mechanism": re.sub(r"\s+", " ", mechanism).strip()[:220],
            "named_actors": re.sub(r"\s+", " ", named_actors).strip()[:220],
            "concrete_assets": re.sub(r"\s+", " ", concrete_assets).strip()[:220],
            "next_date_or_window": re.sub(r"\s+", " ", next_window).strip()[:140],
            "implementation_angle": re.sub(r"\s+", " ", implementation_angle).strip()[:260],
            "scip_reasoning": re.sub(r"\s+", " ", scip_reasoning).strip()[:320],
            "opportunity_strength": re.sub(r"\s+", " ", opportunity_strength).strip()[:40],
            "recommended_zirp_use": re.sub(r"\s+", " ", recommended_zirp_use).strip()[:120],
            "briefing_sentence": re.sub(r"\s+", " ", briefing_sentence).strip()[:280],
            "url": url,
        })

    score_sums: dict[str, list[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])
    for _, row in events_df.iterrows():
        member = str(row.get("mitglied", "")).strip()
        if member not in profiles:
            continue
        title = str(row.get("titel", ""))
        snippet = str(row.get("snippet", ""))
        if is_low_value_event(title, snippet):
            continue
        evidence_weight = evidence_weight_for_level(evidence_level_for_value(row))
        if evidence_weight <= 0:
            continue
        if not classify_clusterable_topic(" ".join([
            title,
            snippet,
            str(row.get("mechanism", "")),
            str(row.get("concrete_assets", "")),
            str(row.get("implementation_angle", "")),
            str(row.get("scip_reasoning", "")),
            str(row.get("opportunity_strength", "")),
            str(row.get("recommended_zirp_use", "")),
            str(row.get("briefing_sentence", "")),
            str(row.get("hauptsektion", "")),
            str(row.get("themenfelder", "")),
            member,
        ])):
            continue
        recency_weight = float(row.get("recent_priority_weight", 1.0) or 1.0) * evidence_weight
        score_sums[member][0] += float(row.get("decision_score", 0) or 0) * recency_weight
        score_sums[member][1] += float(row.get("score", 0) or 0) * recency_weight
        score_sums[member][2] += recency_weight
    for member, (decision_sum, score_sum, weighted_n) in score_sums.items():
        if weighted_n:
            profiles[member].avg_decision_score = decision_sum / weighted_n
            profiles[member].avg_event_score = score_sum / weighted_n

    return profiles


def topic_weight(topic: str) -> float:
    cid = plain_cluster_id(topic)
    if cid in CLUSTER_TOPIC_LOOKUP:
        return PRIORITY_CLUSTER_WEIGHTS.get(cid, 6.5)
    return THEME_WEIGHTS.get(topic, 1.2 if topic in BROAD_FIELD_TOPICS else 2.0)


ROLE_COMPLEMENTARITY = {
    frozenset(("problem_owner", "knowledge_actor")): 10.0,
    frozenset(("problem_owner", "technical_or_domain_expert")): 10.0,
    frozenset(("problem_owner", "implementation_actor")): 12.0,
    frozenset(("implementation_actor", "knowledge_actor")): 14.0,
    frozenset(("implementation_actor", "technical_or_domain_expert")): 14.0,
    frozenset(("implementation_actor", "resource_actor")): 10.0,
    frozenset(("knowledge_actor", "resource_actor")): 5.0,
    frozenset(("technical_or_domain_expert", "resource_actor")): 5.0,
    frozenset(("problem_owner", "resource_actor")): 4.0,
    frozenset(("visibility_actor", "implementation_actor")): 4.0,
    frozenset(("visibility_actor", "knowledge_actor")): 2.0,
}

APPLICATION_CONTEXT_ROLES = {
    "problem_owner",
    "implementation_actor",
}

CAPABILITY_ROLES = {
    "knowledge_actor",
    "technical_or_domain_expert",
    "resource_actor",
}


def role_complementarity_score(member_a: str, member_b: str) -> float:
    roles_a = infer_member_roles(member_a)
    roles_b = infer_member_roles(member_b)
    score = 0.0
    for role_a in roles_a:
        for role_b in roles_b:
            score = max(score, ROLE_COMPLEMENTARITY.get(frozenset((role_a, role_b)), 0.0))
    return score


def has_application_capability_bridge(member_a: str, member_b: str) -> bool:
    roles_a = infer_member_roles(member_a)
    roles_b = infer_member_roles(member_b)
    return (
        bool((roles_a & APPLICATION_CONTEXT_ROLES) and (roles_b & CAPABILITY_ROLES))
        or bool((roles_b & APPLICATION_CONTEXT_ROLES) and (roles_a & CAPABILITY_ROLES))
    )


def pattern_has_application_capability_bridge(pattern: MeetingPattern) -> bool:
    if len(pattern.members) < 2:
        return False
    for member_a, member_b in combinations(pattern.members, 2):
        if has_application_capability_bridge(member_a, member_b):
            return True
    return False


def pattern_role_complementarity_score(pattern: MeetingPattern) -> float:
    if len(pattern.members) < 2:
        return 0.0
    return max(
        role_complementarity_score(member_a, member_b)
        for member_a, member_b in combinations(pattern.members, 2)
    )


def complementarity_text_for_profiles(a: MemberProfile, b: MemberProfile) -> str:
    signal_text = " ".join(
        " ".join(str(signal.get(key, "")) for key in [
            "title",
            "snippet",
            "mechanism",
            "concrete_assets",
            "implementation_angle",
            "scip_reasoning",
            "recommended_zirp_use",
        ])
        for profile in [a, b]
        for signal in profile.signals[:4]
    )
    return " ".join([a.name, b.name, signal_text])


def abstract_overlap_penalty(shared: tuple[str, ...] | list[str], bridge: tuple[str, ...] | list[str], pattern_text: str) -> float:
    shared_topics = list(shared or [])
    bridge_topics = list(bridge or [])
    concrete = any(cluster_info(topic) for topic in shared_topics + bridge_topics)
    has_problem = has_clear_convening_problem_text(pattern_text)

    penalty = 0.0
    if shared_topics and not concrete:
        penalty += 8.0
    if not has_problem:
        penalty += 10.0
    return penalty


def has_valid_complementarity_sentence(pattern: MeetingPattern) -> bool:
    roles_by_member = {member: infer_member_roles(member) for member in pattern.members}
    all_roles: set[str] = set()
    for roles in roles_by_member.values():
        all_roles |= roles

    if "implementation_actor" in all_roles and ("knowledge_actor" in all_roles or "technical_or_domain_expert" in all_roles):
        return True
    if "problem_owner" in all_roles and ("knowledge_actor" in all_roles or "technical_or_domain_expert" in all_roles):
        return True
    if "problem_owner" in all_roles and "implementation_actor" in all_roles:
        return True
    if "resource_actor" in all_roles and ("implementation_actor" in all_roles or "problem_owner" in all_roles):
        return True
    if "visibility_actor" in all_roles:
        return "implementation_actor" in all_roles or "problem_owner" in all_roles
    return False


def pair_score(a: MemberProfile, b: MemberProfile) -> tuple[float, tuple[str, ...], tuple[str, ...], tuple[str, ...]]:
    topics_a = set(a.topics)
    topics_b = set(b.topics)
    shared = sorted(topics_a & topics_b, key=lambda t: -topic_weight(t))
    union = topics_a | topics_b
    if not union:
        return 0.0, tuple(), tuple(), tuple()
    if not (a.clusters or b.clusters):
        return 0.0, tuple(), tuple(), tuple()

    shared_clusters = [t for t in shared if plain_cluster_id(t) in CLUSTER_TOPIC_IDS]
    fallback_clusters: list[str] = []
    if not shared_clusters:
        fallback_counter = a.clusters + b.clusters
        fallback_clusters = [
            topic for topic, _ in fallback_counter.most_common(2)
            if plain_cluster_id(topic) in CLUSTER_TOPIC_IDS
        ]

    shared_score = sum(
        topic_weight(t) * (1 + min(a.topics[t], b.topics[t]) * (0.18 if plain_cluster_id(t) in CLUSTER_TOPIC_IDS else 0.04))
        for t in shared
        if plain_cluster_id(t) in CLUSTER_TOPIC_IDS
    )
    bridge_topics: set[str] = set()
    complement_score = 0.0
    for ta in topics_a:
        for tb in topics_b:
            pair = frozenset((ta, tb))
            if pair in COMPLEMENTARY_TOPIC_PAIRS and ta != tb:
                complement_score += COMPLEMENTARY_TOPIC_PAIRS[pair]
                bridge_topics.update([ta, tb])

    core_overlap = set(shared_clusters or fallback_clusters)
    momentum = math.sqrt(max(a.event_count, 0) + 1) + math.sqrt(max(b.event_count, 0) + 1)
    quality = (a.avg_decision_score + b.avg_decision_score) / 8
    diversity = 2.5 if a.sections and b.sections and set(a.sections) != set(b.sections) else 0.0
    names = tuple(name for name, _ in (a.names + b.names).most_common(5))

    credibility_bonus = 3.0 if core_overlap else 0.0
    role_bridge_score = role_complementarity_score(a.name, b.name)
    has_application_bridge = has_application_capability_bridge(a.name, b.name)
    score = shared_score + complement_score + momentum + quality + diversity + credibility_bonus
    score += role_bridge_score
    if has_application_bridge:
        score += min(8.0, max(4.0, role_bridge_score * 0.60))
    else:
        score -= 12.0
    score -= abstract_overlap_penalty(
        tuple(shared_clusters or fallback_clusters),
        tuple(sorted(bridge_topics)),
        complementarity_text_for_profiles(a, b),
    )
    if is_broad_media_profile(a) or is_broad_media_profile(b):
        media_topics = core_overlap | (set(bridge_topics) & {"gesellschaft", "kultur", "versorgung_gesundheit", "gesundheit"})
        if media_topics:
            score -= 2.5
        else:
            score -= 8.0
    if profile_strength(a) < 4 or profile_strength(b) < 4:
        score -= 4.0
    if not shared and complement_score < 3:
        score -= 5.0
    score -= academic_support_penalty((a.name, b.name))
    return score, tuple((shared_clusters or fallback_clusters)[:6]), tuple(sorted(bridge_topics)[:6]), names


def cluster_score(members: tuple[MemberProfile, ...]) -> tuple[float, tuple[str, ...], tuple[str, ...], tuple[str, ...]]:
    pair_results = []
    for a, b in combinations(members, 2):
        pair_results.append(pair_score(a, b))
    if not pair_results:
        return 0.0, tuple(), tuple(), tuple()

    pair_scores = [x[0] for x in pair_results]
    shared_counter = Counter(topic for _, shared, _, _ in pair_results for topic in shared)
    bridge_counter = Counter(topic for _, _, bridge, _ in pair_results for topic in bridge)
    names_counter = Counter(name for _, _, _, names in pair_results for name in names)

    avg_pair = sum(pair_scores) / len(pair_scores)
    cohesion = len([s for s in pair_scores if s >= 8]) / len(pair_scores)
    size_bonus = {2: 0.0, 3: 3.0, 4: 4.0}.get(len(members), 0.0)
    active_bonus = sum(1 for m in members if m.event_count > 0)
    score = avg_pair + cohesion * 6 + size_bonus + active_bonus
    score -= academic_support_penalty(tuple(member.name for member in members))

    shared = tuple(topic for topic, _ in shared_counter.most_common(6))
    bridge = tuple(topic for topic, _ in bridge_counter.most_common(6))
    names = tuple(name for name, _ in names_counter.most_common(6))
    return score, shared, bridge, names


def pattern_is_editorially_credible(
        members: tuple[MemberProfile, ...],
        score: float,
        shared: tuple[str, ...],
        bridge: tuple[str, ...],
        *,
        allow_backfill: bool = False,
) -> bool:
    topics = set(shared) | set(bridge)
    min_score = 6 if allow_backfill else 9
    if score < min_score:
        return False
    if not (topics & CLUSTER_TOPIC_ALIASES):
        return False
    min_strong_members = 1 if allow_backfill else 2
    if sum(1 for member in members if profile_strength(member) >= 4) < min_strong_members:
        return False
    if len(members) >= 3 and len(topics & CLUSTER_TOPIC_ALIASES) < 1:
        return False
    broad_media_count = sum(1 for member in members if is_broad_media_profile(member))
    if broad_media_count and not (topics & {"gesellschaft", "kultur", "versorgung_gesundheit", "gesundheit"}):
        return False
    if broad_media_count >= 2:
        return False
    return True


def reason_for_pattern(members: tuple[MemberProfile, ...], shared: tuple[str, ...], bridge: tuple[str, ...]) -> str:
    member_names = ", ".join(m.name for m in members)
    if shared:
        return f"{member_names}: gemeinsamer Gesprächskern bei {', '.join(shared[:4])}."
    if bridge:
        return f"{member_names}: komplementäre Perspektiven zwischen {', '.join(bridge[:4])}."
    return f"{member_names}: exploratives Treffen wegen paralleler Signale im Briefing."


def display_topic(topic: str) -> str:
    info = cluster_info(topic)
    if info:
        return str(info.get("label", plain_cluster_id(topic)))
    labels = {
        "versorgung_gesundheit": "Versorgung und Gesundheit",
        "wirtschaftsentwicklung": "Wirtschafts- und Standortentwicklung",
        "kooperationen": "Kooperation und Transfer",
        "fuehrungswechsel": "Führung und institutionelle Prioritäten",
        "soziales_engagement": "gesellschaftliche Verantwortung",
        "technologie": "Technologie",
        "nachhaltigkeit": "Nachhaltigkeit",
        "wissen": "Wissenstransfer",
        "wirtschaft": "Wirtschaft",
        "gesellschaft": "Gesellschaft",
        "pflege": "Pflege und Gesundheit",
        "gesundheit": "Gesundheit",
        "kultur": "Kultur",
    }
    return labels.get(topic, topic.replace("_", " ").title())


def convening_theme(shared: tuple[str, ...], bridge: tuple[str, ...]) -> str:
    topics = list(shared[:3]) or list(bridge[:3])
    cluster_topics = [topic for topic in topics if cluster_info(topic)]
    if cluster_topics:
        return display_topic(cluster_topics[0])
    if not topics:
        return "Exploratives Mitgliedergespräch zu aktuellen Signalen"
    if "versorgung_gesundheit" in topics or "pflege" in topics or "gesundheit" in topics:
        return "Versorgung, Fachkräfte und institutionelle Verantwortung"
    if "technologie" in topics and "nachhaltigkeit" in topics:
        return "Technologie als Hebel für Transformation und Nachhaltigkeit"
    if "wirtschaftsentwicklung" in topics or "wirtschaft" in topics:
        return "Standortentwicklung, Wettbewerbsfähigkeit und Transformation"
    if "kooperationen" in topics or "wissen" in topics:
        return "Kooperation, Transfer und regionale Innovationsfähigkeit"
    if "soziales_engagement" in topics or "gesellschaft" in topics:
        return "Gesellschaftliche Verantwortung und regionale Wirkung"
    return ", ".join(display_topic(t) for t in topics)


def member_capability_sentence(profile: MemberProfile, focus_topics: tuple[str, ...]) -> str:
    top_topics = [display_topic(t) for t, _ in profile.topics.most_common(4)]
    focus = [display_topic(t) for t in focus_topics if t in profile.topics]
    if focus:
        return f"{profile.name} bringt aktuelle Signale zu {', '.join(focus[:3])} ein."
    if top_topics:
        return f"{profile.name} bringt eine Perspektive aus {', '.join(top_topics[:3])} ein."
    return f"{profile.name} ergänzt die Runde mit einer Mitgliedsperspektive aus dem aktuellen Monitoring."


def editorial_justification_for_pattern(
        members: tuple[MemberProfile, ...],
        shared: tuple[str, ...],
        bridge: tuple[str, ...],
) -> str:
    topics = tuple(dict.fromkeys([*shared, *bridge]))
    question_topics = shared if shared else topics
    readable_topics = ", ".join(display_topic(topic) for topic in topics[:3])
    if not readable_topics:
        readable_topics = "einem gemeinsamen Zukunftsthema"
    strongest = sorted(members, key=profile_strength, reverse=True)
    first = strongest[0]
    second = strongest[1] if len(strongest) > 1 else strongest[0]
    return (
        f"Der Anlass ist prüfbar: {first.name} bringt aktuell {profile_focus(first)} ein, "
        f"{second.name} ergänzt {profile_focus(second)}. "
        f"Für die ZIRP stellt sich die Frage: {zirp_question_for_topics(question_topics)}"
    )


def profile_focus(profile: MemberProfile) -> str:
    topics = [display_topic(topic) for topic, _ in profile.topics.most_common(2)]
    sections = [display_topic(section) for section, _ in profile.sections.most_common(1)]
    focus = topics or sections
    return ", ".join(focus[:2]) if focus else "ein beobachtbares Mitgliedssignal"


def zirp_question_for_topics(topics: tuple[str, ...]) -> str:
    for topic in topics:
        info = cluster_info(topic)
        if info:
            return str(info.get("problem", "Welche konkrete ZIRP-Frage ergibt sich aus diesem Problemraum?"))
    topic_set = set(topics)
    if "versorgung_gesundheit" in topic_set or "pflege" in topic_set or "gesundheit" in topic_set:
        return "Gibt es eine konkrete Versorgungs- oder Qualifizierungsfrage, die einen kleinen Sondierungsraum rechtfertigt?"
    if "technologie" in topic_set and "nachhaltigkeit" in topic_set:
        return "Lässt sich digitale Umsetzung für Nachhaltigkeit, Bau oder Verwaltung an einem konkreten Praxisproblem testen?"
    if "kooperationen" in topic_set or "wissen" in topic_set:
        return "Ist der Transferanlass konkret genug für ein Vorgespräch mit klarer Ergebnisfrage?"
    if "wirtschaft" in topic_set or "wirtschaftsentwicklung" in topic_set:
        return "Entsteht daraus eine Standortfrage mit erkennbarem Bezug zu Rheinland-Pfalz?"
    return "Ist der gemeinsame Problemraum eng genug, um mehr als Beobachtung zu rechtfertigen?"


def recent_signal_lines(members: tuple[MemberProfile, ...], limit: int = 5) -> tuple[str, ...]:
    signals: list[tuple[float, str]] = []
    for profile in members:
        for signal in profile.signals[:4]:
            if is_low_value_event(signal.get("title", ""), signal.get("snippet", "")):
                continue
            if not is_full_convening_evidence(signal):
                continue
            section_bonus = topic_weight(signal.get("section", ""))
            text = f"{profile.name}: {signal.get('title', '')}"
            snippet = signal.get("snippet", "")
            if snippet:
                text += f"  {snippet}"
            mechanism = signal.get("mechanism", "")
            if mechanism:
                text += f"  Mechanismus: {mechanism}"
            assets = signal.get("concrete_assets", "")
            if assets:
                text += f"  Asset: {assets}"
            window = signal.get("next_date_or_window", "")
            if window:
                text += f"  Zeitfenster: {window}"
            angle = signal.get("implementation_angle", "")
            if angle:
                text += f"  Opportunity: {angle}"
            strength = signal.get("opportunity_strength", "")
            recommended_use = signal.get("recommended_zirp_use", "")
            briefing = signal.get("briefing_sentence", "")
            if strength:
                text += f"  Opportunity Strength: {strength}"
            if recommended_use:
                text += f"  ZIRP-Nutzung: {recommended_use}"
            if briefing:
                text += f"  Briefing-Satz: {briefing}"
            signals.append((section_bonus + profile.avg_decision_score / 10, text))
    signals.sort(key=lambda x: -x[0])
    cleaned = []
    seen = set()
    for _, text in signals:
        key = text.lower()
        if key in seen:
            continue
        cleaned.append(text)
        seen.add(key)
        if len(cleaned) >= limit:
            break
    return tuple(cleaned)


def suggested_format_for_pattern(size: int, shared: tuple[str, ...], bridge: tuple[str, ...]) -> str:
    topics = set(shared) | set(bridge)
    if "versorgung_gesundheit" in topics or "pflege" in topics or "gesundheit" in topics:
        return "ZIRPzoom mit Expertinnen- und Praxisperspektive"
    if "wirtschaftsentwicklung" in topics and ("technologie" in topics or "nachhaltigkeit" in topics):
        return "Strategischer Mitglieder-Roundtable zu Standort und Transformation"
    if size == 2:
        return "bilaterales Sondierungsgespräch zur Vorbereitung eines Formats"
    if "kooperationen" in topics or "wissen" in topics or "technologie" in topics:
        return "Workshop mit Projekt- und Transferfokus"
    return "kuratierter Mitglieder-Roundtable mit klarer Leitfrage"


def possible_agenda_for_pattern(theme: str, shared: tuple[str, ...], bridge: tuple[str, ...]) -> str:
    topics = set(shared) | set(bridge)
    shared_set = set(shared)
    if ("technologie" in shared_set or "wissen" in shared_set or "kooperationen" in shared_set) and "versorgung_gesundheit" not in shared_set:
        return "Welcher Transferanlass ist konkret genug für ein Sondierungsgespräch?"
    if "versorgung_gesundheit" in shared_set or "pflege" in shared_set or "gesundheit" in shared_set:
        return "Welche konkrete Versorgungs- oder Qualifizierungsfrage lässt sich aus den aktuellen Signalen ableiten?"
    if "technologie" in topics and "nachhaltigkeit" in topics:
        return "Wo kann digitale Umsetzung in Bau, Industrie oder Verwaltung konkret beschleunigt werden?"
    if "wirtschaftsentwicklung" in topics or "wirtschaft" in topics:
        return "Welche Standortfrage ergibt sich aus den aktuellen Mitgliedssignalen, und ist sie für Rheinland-Pfalz handlungsrelevant?"
    if "kooperationen" in topics or "wissen" in topics:
        return "Welcher Transferanlass ist konkret genug für ein Sondierungsgespräch?"
    return f"Welche belastbare ZIRP-Frage ergibt sich aus dem Themenfeld {theme}?"


def rlp_fit_for_pattern(members: tuple[str, ...], shared: tuple[str, ...], bridge: tuple[str, ...], signals: tuple[str, ...]) -> str:
    haystack = " ".join([*members, *shared, *bridge, *signals]).lower()
    rlp_terms = [
        "rheinland-pfalz", "rlp", "mainz", "trier", "kaiserslautern", "bingen",
        "ludwigshafen", "koblenz", "speyer", "pfalz", "suedwest", "südwest",
    ]
    if any(term in haystack for term in rlp_terms):
        return "hoch"
    if len(members) >= 2:
        return "mittel"
    return "unklar"


def maturity_level_for_label(empfehlung: str) -> str:
    if empfehlung.startswith("DO"):
        return "Level 3 - Sondieren"
    if empfehlung.startswith("WATCH"):
        return "Level 2 - Beobachten"
    if empfehlung.startswith("MAYBE"):
        return "Level 1 - Reframen"
    return "Level 0 - Nicht weiterverfolgen"


def convening_format_for_label(label: str, suggested_format: str) -> str:
    if label.startswith("DO"):
        return "Sondierungsgespräch"
    if label.startswith("WATCH"):
        return suggested_format
    if label.startswith("MAYBE"):
        return suggested_format
    if label.startswith("DROP"):
        return "nicht weiterverfolgen"
    return suggested_format


def missing_actors_for_topics(shared: tuple[str, ...], bridge: tuple[str, ...]) -> str:
    shared_set = set(shared)
    topics = shared_set or (set(shared) | set(bridge))
    actors = []
    if "versorgung_gesundheit" in topics or "pflege" in topics or "gesundheit" in topics:
        actors.extend(["Gesundheitsakteur", "Kommune oder Versorgungsplanung"])
    if "technologie" in topics or "wirtschaftsentwicklung" in topics:
        actors.extend(["Mittelstand", "Wirtschaftsfoerderung"])
    if "nachhaltigkeit" in topics:
        actors.extend(["Kommune", "Ministerium oder Fachverwaltung"])
    if "wissen" in topics or "kooperationen" in topics:
        actors.extend(["Transferstelle", "Kammer oder Praxispartner"])
    return ", ".join(dict.fromkeys(actors[:4])) or "noch zu klaeren"


def next_action_for_label(label: str, members: tuple[str, ...], agenda: str) -> str:
    if label.startswith("DO"):
        return f"ZIRP sollte mit {members[0]} und {members[1]} ein 30-minütiges Sondierungsgespräch zur Frage '{agenda}' prüfen."
    if label.startswith("WATCH"):
        return "Noch kein Format ansetzen; zwei Wochen weiter beobachten und auf ein zweites konkretes Umsetzungssignal warten."
    if label.startswith("MAYBE"):
        return "Nicht als Treffen planen, sondern die Fragestellung enger fassen und fehlende Praxisakteure identifizieren."
    return "Nicht weiterverfolgen, solange kein klarer RLP- oder Handlungsbezug entsteht."


def recommendation_limit_for_label(label: str) -> str:
    if label.startswith("DO"):
        return "Noch kein öffentliches Format planen; zuerst nur einen kurzen Vorkontakt mit klarer Ergebnisfrage."
    if label.startswith("WATCH"):
        return "Noch kein Format ansetzen und nicht als feste Mitglieder-Runde uebernehmen."
    if label.startswith("MAYBE"):
        return "Erst reframen, bevor Mitglieder angesprochen werden."
    return "Nicht weiterverfolgen, solange kein konkreter ZIRP-Handlungswert belegt ist."


def signal_member_name(signal: str) -> str:
    return str(signal).split(":", 1)[0].strip()


def distinct_signal_member_count(signals: tuple[str, ...]) -> int:
    return len({signal_member_name(signal) for signal in signals if signal_member_name(signal)})


def _normalized_pattern_member_set(pattern: MeetingPattern) -> set[str]:
    return {normalize_search_text(member) for member in pattern.members if normalize_search_text(member)}


def _normalized_signal_member_set(signals: tuple[str, ...]) -> set[str]:
    return {normalize_search_text(signal_member_name(signal)) for signal in signals if normalize_search_text(signal_member_name(signal))}


def pair_signal_members(pattern: MeetingPattern) -> set[str]:
    return _normalized_signal_member_set(pattern.recent_signals) & _normalized_pattern_member_set(pattern)


def pair_evidence_balance(pattern: MeetingPattern) -> str:
    matched_members = pair_signal_members(pattern)
    text = normalize_search_text(" ".join([
        pattern.convening_theme,
        pattern.why_these_members,
        pattern.editorial_justification,
        pattern.possible_agenda,
        " ".join(pattern.recent_signals),
    ]))
    macro_hits = normalized_hit_count(MACRO_TERMS | MARKET_NOISE_TERMS, text)
    raw_signal_text = normalize_search_text(" ".join(pattern.recent_signals))
    market_report_hits = normalized_hit_count(MARKET_REPORT_TERMS, raw_signal_text)
    problem_score = weighted_problem_hit_score(text)["score"]
    action_score = weighted_action_hit_score(text)["score"]
    implementation_hits = normalized_hit_count(
        {"pilot", "modellprojekt", "testfeld", "implementierung", "betrieb", "praxispartner"},
        raw_signal_text,
    )

    # Macro evidence can mention both actors and still be weak for a convening.
    # Treat market-report pairings as briefing material unless the raw signals contain a concrete implementation anchor.
    if market_report_hits >= 2 and implementation_hits == 0:
        return "macro_only"
    if macro_hits >= 2 and (action_score < 2 or problem_score < 1):
        return "macro_only"

    if len(matched_members) >= 2:
        return "balanced"

    if len(matched_members) == 1:
        return "one_sided"

    return "indirect_only"


def pair_evidence_balance_bonus(pattern: MeetingPattern) -> float:
    return 6.0 if pair_evidence_balance(pattern) == "balanced" else 0.0


def pair_evidence_balance_penalty(pattern: MeetingPattern) -> float:
    balance = pair_evidence_balance(pattern)
    if balance == "one_sided":
        return 6.0
    if balance == "macro_only":
        return 12.0
    if balance == "indirect_only":
        return 10.0
    return 0.0


def strong_problem_owner_capability_exception(pattern: MeetingPattern) -> bool:
    matched_members = pair_signal_members(pattern)
    if len(matched_members) != 1:
        return False

    problem_signal_member = next(iter(matched_members))
    roles_by_member = {
        normalize_search_text(member): infer_member_roles(member)
        for member in pattern.members
    }
    signal_roles = roles_by_member.get(problem_signal_member, set())
    if not (signal_roles & {"problem_owner", "decision_maker"}):
        return False

    counterpart_roles = set().union(*[
        roles
        for member, roles in roles_by_member.items()
        if member != problem_signal_member
    ]) if roles_by_member else set()
    has_capability_actor = bool(counterpart_roles & {
        "knowledge_actor",
        "technical_or_domain_expert",
        "implementation_actor",
        "content_expert",
    })
    if not has_capability_actor:
        return False

    text = normalize_search_text(" ".join([
        pattern.convening_theme,
        pattern.why_these_members,
        pattern.editorial_justification,
        pattern.possible_agenda,
        " ".join(pattern.recent_signals),
    ]))
    problem_hits = weighted_problem_hit_score(text)
    return problem_hits["static"] >= 1 and problem_hits["score"] >= 1


def balance_allows_sondieren(pattern: MeetingPattern) -> bool:
    return pair_evidence_balance(pattern) == "balanced" or strong_problem_owner_capability_exception(pattern)


def has_non_generic_zirp_question(shared: tuple[str, ...], bridge: tuple[str, ...]) -> bool:
    topics = set(shared)
    if any(cluster_info(topic) for topic in topics):
        return True
    if topics & {"versorgung_gesundheit", "pflege", "gesundheit", "technologie", "nachhaltigkeit", "wissen", "kooperationen", "wirtschaftsentwicklung", "wirtschaft"}:
        return True
    return bool(set(shared) & set(bridge))


def why_now_for_pattern(members: tuple[str, ...], signals: tuple[str, ...], label: str) -> str:
    useful = [s.split("  ")[0].strip() for s in signals[:2]]
    if useful:
        basis = "; ".join(useful)
        return f"Aktuelle Auslöser sind: {basis}. Das Zeitfenster ist relevant, weil sich daraus eine konkrete ZIRP-Frage ableiten lässt."
    return "Der Anlass ist noch schwach; die Empfehlung beruht derzeit vor allem auf thematischer Naehe."


def actor_judgment_role(member: str) -> str:
    lower = repair_mojibake(member).lower()
    if "iqib" in lower:
        return "bringt angewandte Innovationsforschung und Resilienzperspektive ein"
    if "universit" in lower and "trier" in lower:
        return "bringt wissenschaftliche, ethische und gesellschaftliche Orientierung ein"
    if "hochschule mainz" in lower:
        return "bringt anwendungsnahen Transfer und Projektzugang ein"
    if "technische hochschule" in lower or "th bingen" in lower:
        return "bringt technische Transferkompetenz und Nachwuchsperspektive ein"
    if "zahnen" in lower:
        return "bringt digitale Anlagen-, Infrastruktur- und Prozesspraxis ein"
    if "handwerkskammer" in lower:
        return "bringt Zugang zu Betrieben und mittelständischer Umsetzungspraxis ein"
    if "swr" in lower:
        return "bringt regionale Sichtbarkeit und Problemwahrnehmung ein, aber nur begrenzte Umsetzungskapazität"
    role = actor_role(member)
    if role == "academic_support":
        return "bringt Wissens-, Forschungs- und Transferperspektive ein"
    if role == "profit_anchor":
        return "bringt Praxis-, Markt- oder Anwendungsperspektive ein"
    if role == "implementation_anchor":
        return "bringt Zugang zu Umsetzung, Fachpraxis oder Mitgliedern ein"
    return "ist vor allem Kontextakteur; die aktive Convening-Rolle muss validiert werden"


def concrete_joint_action_for_pattern(pattern: MeetingPattern) -> str:
    actors = " | ".join(pattern.members).lower()
    topics = set(pattern.shared_topics) | set(pattern.bridge_topics) | set(pattern.concrete_clusters)
    if "zahnen" in actors and "hochschule mainz" in actors:
        return "90-Tage-Testfeld: eine digitale Infrastruktur- oder Prozessanwendung unter realen Bedingungen prüfen."
    if "iqib" in actors and ("universit" in actors and "trier" in actors):
        return "Sondierung: kommunales Resilienzformat mit Ethik-, Akzeptanz- und Umsetzungskriterien schärfen."
    if "iqib" in actors and "hochschule mainz" in actors:
        return "Transfer-Sprint: aus Resilienz- und Transferwissen eine testbare kommunale Anwendungsskizze entwickeln."
    if ("bundesagentur" in actors or "agentur fuer arbeit" in actors or "agentur für arbeit" in actors) and ("universit" in actors or "hochschule" in actors):
        return "Vorkontakt: eine konkrete RLP-Qualifizierungs-, Weiterbildungs- oder Arbeitsmarktfrage klaeren."
    if "handwerkskammer" in actors and ("universit" in actors and "trier" in actors):
        return "Praxischeck: digitale Qualifizierungs- oder Akzeptanzfragen in einem mittelständischen Anwendungskontext prüfen."
    if "handwerkskammer" in actors and ("technische hochschule" in actors or "th bingen" in actors):
        return "Praxischeck: ein digitales Werkzeug für betriebliche Abläufe in einem Handwerkscluster testen."
    if "zahnen" in actors and "handwerkskammer" in actors:
        return "Praxisworkshop: digitale Prozesspraxis auf übertragbare Mittelstandsbedarfe prüfen."
    if "landesbank" in actors or "lbbw" in actors or "sparkasse" in actors:
        return "Finanzierungscheck: ein Investitionshemmnis für digitale Infrastruktur in ein prüfbares Förder- oder Finanzierungsszenario übersetzen."
    if "swr" in actors:
        return "Briefing- oder Sichtbarkeitsanlass prüfen; Umsetzungspartner separat validieren."
    if topics & {"mittelstand_digitalisierung_produktivitaet", "technologie"}:
        return "90-Tage-Praxischeck: eine digitale Anwendung für mittelständische Produktivität in einem Betriebskontext testen."
    if topics & {"verwaltungswissen_digitalisierung"}:
        return "Sondierung: eine konkrete Verwaltungsaufgabe als digitales Testfeld definieren."
    if topics & {"cyber_resilienz_kritische_infrastruktur", "energie_versorgungssicherheit"}:
        return "Testfeld: eine Resilienz- oder Infrastrukturfrage mit einem konkreten Betreiberkontext prüfen."
    if topics & {"finanzierung_foerderung_transformation", "standort_investition_transformation"}:
        return "Sondierung: ein konkretes Investitionshemmnis identifizieren und als 90-Tage-Lösungspfad prüfen."
    agenda = repair_mojibake(pattern.possible_agenda).strip()
    if agenda and not any(term in agenda.lower() for term in ["bau, industrie oder verwaltung", "bau, industrie", "industrie oder verwaltung"]):
        return f"Sondierungsgespräch mit Ergebnisfrage: {agenda}"
    return "Erst eine konkrete Pilotfrage formulieren; noch kein Format ansetzen."


def critical_judgment_for_pattern(
        pattern: MeetingPattern,
        labels: dict[str, Any],
        evidence_count: int,
        distinct_evidence_members: int,
        rlp_fit: str,
        common_problem_quality: str,
) -> dict[str, str]:
    ordered = ordered_members_for_display(pattern.members)
    first = ordered[0] if ordered else ""
    second = ordered[1] if len(ordered) > 1 else ""
    counts = actor_role_counts(pattern.members)
    why_actor = f"{first} {actor_judgment_role(first)}." if first else "Akteursrolle noch nicht ausreichend geklärt."
    why_counterpart = f"{second} {actor_judgment_role(second)}." if second else "Gegenpart noch nicht ausreichend geklärt."
    balance = pair_evidence_balance(pattern)
    balance_exception = strong_problem_owner_capability_exception(pattern)
    if balance == "balanced" and evidence_count >= 4 and distinct_evidence_members >= 2:
        evidence_quality = "hoch: mehrere aktuelle Signale aus mindestens zwei Akteurskontexten."
    elif balance == "balanced" and evidence_count >= 2:
        evidence_quality = "mittel: mehrere Signale aus beiden Akteurskontexten, aber Rollen und Bedarf sollten validiert werden."
    elif balance == "one_sided":
        evidence_quality = "mittel bis niedrig: die Akteurslogik ist plausibel, aber der aktuelle Anlass kommt vor allem von einem Akteur."
    elif balance == "macro_only":
        evidence_quality = "niedrig: die Signale stuetzen eher ein Briefing als ein Convening."
    elif balance == "indirect_only":
        evidence_quality = "niedrig: die Verbindung ist indirekt und noch nicht durch aktuelle Paar-Evidenz belegt."
    else:
        evidence_quality = "niedrig: bisher zu wenig unabhängige Evidenz für eine direkte Convening-Empfehlung."
    source_only_risk = counts["context_actor"] and not counts["academic_support"]
    if balance == "macro_only":
        sufficiency = "ausreichend für ein internes Briefing, aber nicht für ein Convening."
        action_decision = "brief"
    elif labels["empfehlung"].startswith("DO") and not source_only_risk and evidence_count >= 2 and distinct_evidence_members >= 2 and (balance == "balanced" or balance_exception):
        sufficiency = "ausreichend für ein kuratiertes Sondierungsgespräch."
        action_decision = "sondieren"
    elif source_only_risk:
        sufficiency = "nicht ausreichend für direktes Convening; Kontextakteur oder Quelle zuerst validieren."
        action_decision = "brief"
    elif evidence_count >= 2:
        sufficiency = "ausreichend für Beobachtung oder Reframing, noch nicht für ein Top-Match."
        action_decision = "watch"
    else:
        sufficiency = "nicht ausreichend; weiteres Signal oder konkreter Bedarf fehlt."
        action_decision = "watch"
    if labels["empfehlung"].startswith("DROP"):
        action_decision = "ignore"
    elif labels["empfehlung"].startswith("MAYBE"):
        action_decision = "reframe"
    consequence = shared_tension_for_pattern(pattern.shared_topics, pattern.bridge_topics)
    uncertainty_parts = []
    if source_only_risk:
        uncertainty_parts.append("ob der Kontextakteur eine aktive Rolle hat")
    if balance == "one_sided":
        uncertainty_parts.append("ob die zweite Seite den Bedarf aktuell bestaetigt")
    elif balance == "indirect_only":
        uncertainty_parts.append("ob aus indirekter Evidenz ein konkreter RLP-Anwendungsanker entsteht")
    elif balance == "macro_only":
        uncertainty_parts.append("ob die Makro-Signale ueberhaupt eine konkrete Convening-Frage tragen")
    elif distinct_evidence_members < 2:
        uncertainty_parts.append("ob beide Seiten den Bedarf aktuell bestätigen")
    if common_problem_quality == "schwach":
        uncertainty_parts.append("ob die gemeinsame Problemstellung eng genug ist")
    if rlp_fit == "niedrig":
        uncertainty_parts.append("ob der RLP-Bezug handlungsrelevant ist")
    uncertainty = "; ".join(uncertainty_parts) if uncertainty_parts else "Rollen, Bedarf und Pilotfrage sind plausibel, müssen aber im Vorkontakt bestätigt werden."
    reason = f"Entscheidung: {action_decision}, weil Evidenzqualität {evidence_quality.split(':', 1)[0]} ist und die Suffizienz {sufficiency}"
    return {
        "why_this_actor": why_actor,
        "why_this_counterpart": why_counterpart,
        "why_now": why_now_for_pattern(pattern.members, pattern.recent_signals, labels["empfehlung"]),
        "concrete_joint_action": concrete_joint_action_for_pattern(pattern),
        "evidence_quality": evidence_quality,
        "sufficiency": sufficiency,
        "consequence": consequence,
        "uncertainty": uncertainty,
        "action_decision": action_decision,
        "critical_decision_reason": reason,
    }


def shared_tension_for_pattern(shared: tuple[str, ...], bridge: tuple[str, ...]) -> str:
    topics = set(shared) | set(bridge)
    shared_set = set(shared)
    if ("technologie" in shared_set or "wissen" in shared_set or "kooperationen" in shared_set) and "versorgung_gesundheit" not in shared_set:
        return "Welche Transferfrage ist konkret genug, damit aus Hochschul- oder Netzwerkaktivität ein ZIRP-Gespräch wird?"
    if "versorgung_gesundheit" in shared_set or "pflege" in shared_set or "gesundheit" in shared_set:
        return "Wie werden Qualifizierung, Versorgungskapazitaet und institutionelle Verantwortung praktisch zusammengebracht?"
    if "technologie" in topics and "nachhaltigkeit" in topics:
        return "Wie werden digitale Werkzeuge aus Pilot- oder Fachlogik in konkrete Umsetzung für Bau, Verwaltung oder Mittelstand überführt?"
    if "kooperationen" in topics or "wissen" in topics:
        return "Welche Transferfrage ist konkret genug, damit aus Hochschul- oder Netzwerkaktivität ein ZIRP-Gespräch wird?"
    return "Die gemeinsame Problemstellung ist noch nicht eng genug; ZIRP sollte sie vor einem Format schärfen."


def build_convening_fields(
        members: tuple[MemberProfile, ...],
        shared: tuple[str, ...],
        bridge: tuple[str, ...],
) -> tuple[str, str, str, tuple[str, ...], str, str]:
    focus_topics = tuple(dict.fromkeys([*shared, *bridge]))
    theme = convening_theme(shared, bridge)
    why = " ".join(member_capability_sentence(member, focus_topics) for member in members)
    editorial = editorial_justification_for_pattern(members, shared, bridge)
    signals = recent_signal_lines(members)
    fmt = suggested_format_for_pattern(len(members), shared, bridge)
    agenda = possible_agenda_for_pattern(theme, shared, bridge)
    return theme, why, editorial, signals, fmt, agenda


def _clean_cell(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _float_cell(value: Any, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def _int_from_candidate_id(candidate_id: str, fallback: int) -> int:
    match = re.search(r"(\d+)", str(candidate_id or ""))
    if match:
        return int(match.group(1))
    return fallback


def _static_support_score_from_row(row_dict: dict[str, Any]) -> float:
    for column in ["frontier_score", "card_score", "final_scip_score", "final_support_score", "legacy_score"]:
        explicit = _float_cell(row_dict.get(column), math.nan)
        if not math.isnan(explicit) and explicit > 0:
            return explicit
    for column in ["role_subrole_balanced_score", "subrole_balanced_score", "role_balanced_score"]:
        quota_score = _float_cell(row_dict.get(column), math.nan)
        if not math.isnan(quota_score) and quota_score > 0:
            return quota_score
    base_score = _float_cell(row_dict.get("base_support_score"), math.nan)
    if not math.isnan(base_score) and base_score > 0:
        return base_score
    positive_columns = [
        "shared_concrete_clusters",
        "complementary_topic_pairs",
        "momentum_from_event_count",
        "avg_decision_quality",
        "section_diversity",
        "role_complementarity",
        "application_capability_bridge",
    ]
    penalty_columns = [
        "broad_media_penalty",
        "weak_profile_penalty",
        "academic_only_penalty",
        "abstract_overlap_penalty",
    ]
    positive = sum(_float_cell(row_dict.get(column), 0.0) for column in positive_columns)
    penalty = sum(_float_cell(row_dict.get(column), 0.0) for column in penalty_columns)
    return max(0.0, positive - penalty)


def _static_cluster_keyword_map(path: Path) -> dict[str, set[str]]:
    if not path.exists():
        return {}
    try:
        df = pd.read_excel(path, sheet_name=STATIC_CLUSTER_SHEET)
    except Exception:
        return {}
    result: dict[str, set[str]] = {}
    for _, row in df.iterrows():
        cluster = _clean_cell(row.get("concrete_cluster"))
        keywords = _clean_cell(row.get("signal_keywords"))
        if not cluster:
            continue
        result[cluster] = {
            normalize_search_text(part)
            for part in re.split(r"[,;]", keywords)
            if normalize_search_text(part)
        }
    return result



def _read_excel_sheet_optional(path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        if path.exists():
            return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as exc:
        if "Worksheet named" not in str(exc):
            print(f"Excel static loader: could not read {sheet_name} from {path}: {exc}")
    return pd.DataFrame()


def _read_excel_table_optional(path: Path, sheet_name: str, required_columns: set[str]) -> pd.DataFrame:
    """Read workbook tables that may have title rows before the header."""
    if not path.exists():
        return pd.DataFrame()
    try:
        direct = pd.read_excel(path, sheet_name=sheet_name)
        if required_columns.issubset({str(column).strip() for column in direct.columns}):
            return direct.dropna(how="all").copy()
    except Exception:
        pass
    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception as exc:
        if "Worksheet named" not in str(exc):
            print(f"Excel static loader: could not read {sheet_name} from {path}: {exc}")
        return pd.DataFrame()
    for idx in range(min(len(raw), 12)):
        headers = [_clean_cell(value) for value in raw.iloc[idx].tolist()]
        if required_columns.issubset(set(headers)):
            table = raw.iloc[idx + 1 :].copy()
            table.columns = headers
            table = table.dropna(how="all")
            return table.loc[:, [column for column in table.columns if column]].copy()
    return pd.DataFrame()


def _candidate_rows_without_removed_members(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or not {"member_a", "member_b"}.issubset(df.columns):
        return df
    keep_mask = []
    for _, row in df.iterrows():
        a = normalize_search_text(row.get("member_a", ""))
        b = normalize_search_text(row.get("member_b", ""))
        keep_mask.append(a not in STATIC_REMOVED_MEMBERS and b not in STATIC_REMOVED_MEMBERS)
    return df.loc[keep_mask].copy()


def _unique_static_candidate_ids(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure candidate_id/scip_variable are unique after optional patch merging."""
    if df.empty:
        return df
    df = df.copy().reset_index(drop=True)
    if "candidate_id" not in df.columns:
        df["candidate_id"] = ""
    if "scip_variable" not in df.columns:
        df["scip_variable"] = ""
    candidate_ids = df["candidate_id"].astype(str).fillna("")
    scip_vars = df["scip_variable"].astype(str).fillna("")
    needs_renumber = candidate_ids.duplicated().any() or scip_vars.duplicated().any() or (candidate_ids.str.strip() == "").any()
    if needs_renumber:
        df["excel_candidate_id"] = candidate_ids
        df["candidate_id"] = [f"C{i:04d}" for i in range(1, len(df) + 1)]
        df["scip_variable"] = [f"x_{i:04d}" for i in range(1, len(df) + 1)]
    return df


def _normalize_full_pairing_universe_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Make the 5,253-pair universe compatible with the legacy static candidate pipeline."""
    if df.empty:
        return df
    df = df.copy().reset_index(drop=True)
    if "candidate_id" not in df.columns:
        df["candidate_id"] = df.get("pair_id", pd.Series(dtype=object))
    if "scip_variable" not in df.columns:
        df["scip_variable"] = [f"x_pair_{i:04d}" for i in range(1, len(df) + 1)]
    if "concrete_cluster" not in df.columns and "top_cluster" in df.columns:
        df["concrete_cluster"] = df["top_cluster"]
    if "cluster_label" not in df.columns and "top_cluster_label" in df.columns:
        df["cluster_label"] = df["top_cluster_label"]
    if "static_cluster_problem" not in df.columns and "top_static_problem" in df.columns:
        df["static_cluster_problem"] = df["top_static_problem"]
    if "final_support_score" not in df.columns and "final_scip_score" in df.columns:
        df["final_support_score"] = df["final_scip_score"]
    if "geo_proximity_norm" not in df.columns and "geo_proximity_score" in df.columns:
        df["geo_proximity_norm"] = pd.to_numeric(df["geo_proximity_score"], errors="coerce").fillna(0) / 100.0
    if "opportunity_key" not in df.columns:
        pair_id = df.get("pair_id", df["candidate_id"]).astype(str)
        cluster = df.get("concrete_cluster", "").astype(str) if "concrete_cluster" in df.columns else ""
        df["opportunity_key"] = pair_id + "__" + cluster
    if "reasoning" not in df.columns:
        status = df.get("candidate_status", "").astype(str) if "candidate_status" in df.columns else ""
        reason = df.get("filter_reason", "").astype(str) if "filter_reason" in df.columns else ""
        geo = df.get("geo_score_reason", "").astype(str) if "geo_score_reason" in df.columns else ""
        df["reasoning"] = (status + " | " + reason + " | " + geo).str.strip(" |")
    if "weak_evidence_flag" not in df.columns:
        gate_source = df["evidence_gate_multiplier"] if "evidence_gate_multiplier" in df.columns else pd.Series([1.0] * len(df))
        gate = pd.to_numeric(gate_source, errors="coerce").fillna(1.0)
        df["weak_evidence_flag"] = (gate < 0.75).astype(int)
    if "low_score_flag" not in df.columns:
        score_source = df["final_support_score"] if "final_support_score" in df.columns else pd.Series([0.0] * len(df))
        score = pd.to_numeric(score_source, errors="coerce").fillna(0.0)
        status = df["candidate_status"].astype(str).str.lower() if "candidate_status" in df.columns else pd.Series([""] * len(df))
        df["low_score_flag"] = ((score < MIN_FINAL_PAIR_SCORE) | status.str.startswith("low")).astype(int)
    if "academic_pair_flag" not in df.columns and "includes_academic" in df.columns:
        df["academic_pair_flag"] = pd.to_numeric(df["includes_academic"], errors="coerce").fillna(0).astype(int)
    if "academic_academic_flag" not in df.columns and "academic_actor_slots" in df.columns:
        slots = pd.to_numeric(df["academic_actor_slots"], errors="coerce").fillna(0)
        df["academic_academic_flag"] = (slots >= 2).astype(int)
    if "non_academic_pair_flag" not in df.columns and "academic_actor_slots" in df.columns:
        slots = pd.to_numeric(df["academic_actor_slots"], errors="coerce").fillna(0)
        df["non_academic_pair_flag"] = (slots == 0).astype(int)
    return df


def _fill_from_merge_column(df: pd.DataFrame, column: str, suffix: str = "_pairvec") -> pd.DataFrame:
    merge_column = f"{column}{suffix}"
    if merge_column not in df.columns:
        return df
    if column not in df.columns:
        df[column] = df[merge_column]
    else:
        current = df[column]
        missing = current.isna() | current.astype(str).str.strip().eq("")
        df.loc[missing, column] = df.loc[missing, merge_column]
    return df.drop(columns=[merge_column])


def _normalize_nested_vector_space_dataframe(path: Path, df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Map the nested vector-space workbook into the optimizer's candidate schema."""
    if df.empty:
        return df
    df = df.copy().reset_index(drop=True)

    pair_vector = _read_excel_sheet_optional(path, STATIC_PAIR_VECTOR_SHEET)
    if not pair_vector.empty and "pair_id" in df.columns and "pair_id" in pair_vector.columns:
        keep_pair_cols = [
            column
            for column in [
                "pair_id", "member_a_id", "member_b_id", "role_a", "role_b", "subrole_a", "subrole_b",
                "distance_km", "same_geo_region", "geo_proximity_score", "geo_bridge_bonus", "geo_score_reason",
                "shared_cluster_count", "weighted_cluster_overlap", "cluster_similarity_score", "top_cluster",
                "top_cluster_label", "top_static_problem", "role_complementarity_score",
                "subrole_complementarity_score", "old_candidate_max_score", "old_candidate_rows",
                "base_structural_score", "evidence_gate_multiplier", "final_scip_score",
                "academic_actor_slots", "includes_academic", "academic_academic_flag",
                "non_academic_pair_flag", "legacy_candidate_id", "legacy_opportunity_key",
            ]
            if column in pair_vector.columns
        ]
        merged = df.merge(
            pair_vector[keep_pair_cols].drop_duplicates("pair_id", keep="last"),
            on="pair_id",
            how="left",
            suffixes=("", "_pairvec"),
        )
        for column in keep_pair_cols:
            if column != "pair_id":
                merged = _fill_from_merge_column(merged, column)
        df = merged

    cards = _read_excel_sheet_optional(path, STATIC_CARD_SHEET)
    if not cards.empty and "card_id" in df.columns and "card_id" in cards.columns:
        keep_card_cols = [
            column
            for column in [
                "card_id", "topic_vector_id", "topic_label", "opportunity_key", "next_step_type",
                "risk_type", "card_score", "card_explanation",
            ]
            if column in cards.columns
        ]
        merged = df.merge(
            cards[keep_card_cols].drop_duplicates("card_id", keep="last"),
            on="card_id",
            how="left",
            suffixes=("", "_card"),
        )
        for column in keep_card_cols:
            if column != "card_id":
                merged = _fill_from_merge_column(merged, column, suffix="_card")
        df = merged

    if "candidate_id" not in df.columns:
        fallback_id = (
            df["card_id"] if "card_id" in df.columns else
            df["frontier_id"] if "frontier_id" in df.columns else
            df["pair_id"] if "pair_id" in df.columns else
            pd.Series([f"POOL_{i:04d}" for i in range(1, len(df) + 1)])
        )
        df["candidate_id"] = fallback_id
        if "frontier_id" in df.columns:
            missing = df["candidate_id"].isna() | df["candidate_id"].astype(str).str.strip().eq("")
            df.loc[missing, "candidate_id"] = df.loc[missing, "frontier_id"]
        if "pair_id" in df.columns:
            missing = df["candidate_id"].isna() | df["candidate_id"].astype(str).str.strip().eq("")
            df.loc[missing, "candidate_id"] = df.loc[missing, "pair_id"]
    if "scip_variable" not in df.columns:
        df["scip_variable"] = [f"x_pool_{i:04d}" for i in range(1, len(df) + 1)]
    if "concrete_cluster" not in df.columns:
        df["concrete_cluster"] = df.get("best_topic_label", df.get("topic_label", df.get("top_cluster_label", "")))
    if "cluster_label" not in df.columns:
        df["cluster_label"] = df.get("best_topic_label", df.get("topic_label", df.get("top_cluster_label", "")))
    if "opportunity_key" not in df.columns:
        df["opportunity_key"] = df.get("best_opportunity_key", df.get("pair_id", df["candidate_id"]))
    else:
        best = df.get("best_opportunity_key")
        if best is not None:
            missing = df["opportunity_key"].isna() | df["opportunity_key"].astype(str).str.strip().eq("")
            df.loc[missing, "opportunity_key"] = best.loc[missing]
    if "static_cluster_problem" not in df.columns:
        df["static_cluster_problem"] = df.get("top_static_problem", "")
    if "reasoning" not in df.columns:
        parts = []
        for column in ["pool_reason", "card_explanation", "filter_reason", "optimization_status"]:
            if column in df.columns:
                parts.append(df[column].fillna("").astype(str))
        if parts:
            reasoning = parts[0]
            for part in parts[1:]:
                reasoning = reasoning.str.cat(part, sep=" | ")
            df["reasoning"] = reasoning.str.replace(r"( \| )+", " | ", regex=True).str.strip(" |")
        else:
            df["reasoning"] = ""
    if "final_support_score" not in df.columns:
        df["final_support_score"] = df.get("frontier_score", df.get("card_score", df.get("final_scip_score", 0.0)))
    if "final_scip_score" not in df.columns:
        df["final_scip_score"] = df["final_support_score"]
    if "geo_proximity_score" not in df.columns and "geo_fit" in df.columns:
        df["geo_proximity_score"] = df["geo_fit"]
    if "geo_proximity_norm" not in df.columns and "geo_proximity_score" in df.columns:
        df["geo_proximity_norm"] = pd.to_numeric(df["geo_proximity_score"], errors="coerce").fillna(0) / 100.0
    if "role_complementarity_score" not in df.columns and "role_fit" in df.columns:
        df["role_complementarity_score"] = df["role_fit"]
    if "subrole_complementarity_score" not in df.columns and "subrole_fit" in df.columns:
        df["subrole_complementarity_score"] = df["subrole_fit"]
    if "old_candidate_max_score" not in df.columns and "legacy_score" in df.columns:
        df["old_candidate_max_score"] = df["legacy_score"]
    if "legacy_final_support_score" in df.columns and "legacy_score" not in df.columns:
        df["legacy_score"] = df["legacy_final_support_score"]
    if "weak_evidence_flag" not in df.columns:
        df["weak_evidence_flag"] = 0
    if "low_score_flag" not in df.columns:
        score = pd.to_numeric(df.get("final_support_score", 0.0), errors="coerce").fillna(0.0)
        df["low_score_flag"] = (score < MIN_FINAL_PAIR_SCORE).astype(int)
    if "academic_pair_flag" not in df.columns:
        slots = pd.to_numeric(df.get("academic_actor_slots", 0), errors="coerce").fillna(0)
        df["academic_pair_flag"] = (slots > 0).astype(int)
    if "academic_academic_flag" not in df.columns:
        slots = pd.to_numeric(df.get("academic_actor_slots", 0), errors="coerce").fillna(0)
        df["academic_academic_flag"] = (slots >= 2).astype(int)
    if "non_academic_pair_flag" not in df.columns:
        slots = pd.to_numeric(df.get("academic_actor_slots", 0), errors="coerce").fillna(0)
        df["non_academic_pair_flag"] = (slots == 0).astype(int)
    if "filter_reason" not in df.columns:
        df["filter_reason"] = df.get("pool_reason", df.get("optimization_status", ""))
    if "member_pair_key" not in df.columns and {"member_a", "member_b"}.issubset(df.columns):
        df["member_pair_key"] = df.apply(
            lambda row: " | ".join(sorted([_clean_cell(row.get("member_a")), _clean_cell(row.get("member_b"))])),
            axis=1,
        )
    df["candidate_source_sheet"] = sheet_name
    return df


def _merge_candidate_role_balance(path: Path, df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    balance = _read_excel_table_optional(
        path,
        "Candidate_Role_Balance",
        {"candidate_id", "role_balanced_score"},
    )
    if balance.empty:
        return df
    keep_cols = [
        column
        for column in [
            "candidate_id", "role_balanced_score", "pair_role_mass",
            "role_balance_multiplier", "weight_a_1_over_role",
            "weight_b_1_over_role", "subrole_a", "subrole_b",
            "subrole_weight_a", "subrole_weight_b", "pair_subrole_mass",
            "subrole_multiplier", "combined_role_subrole_multiplier",
            "role_subrole_balanced_score", "academic_or_uni_pair",
        ]
        if column in balance.columns
    ]
    balance = balance[keep_cols].copy()
    balance["candidate_id"] = balance["candidate_id"].map(_clean_cell)
    balance = balance[balance["candidate_id"] != ""].drop_duplicates("candidate_id", keep="last")
    merged = df.copy()
    if "candidate_id" in merged.columns:
        merged["candidate_id"] = merged["candidate_id"].map(_clean_cell)
    for column in keep_cols:
        if column != "candidate_id" and column in merged.columns:
            merged = merged.drop(columns=[column])
    return merged.merge(balance, on="candidate_id", how="left")


def _load_static_member_cluster_matrix(paths: list[Path]) -> dict[tuple[str, str], float]:
    relevance: dict[tuple[str, str], float] = {}
    meta_columns = {"member_id", "member_name", "actor_role"}
    for workbook in paths:
        matrix_df = _read_excel_sheet_optional(workbook, STATIC_MEMBER_CLUSTER_MATRIX_SHEET)
        if matrix_df.empty:
            continue
        cluster_columns = [column for column in matrix_df.columns if str(column) not in meta_columns]
        for _, row in matrix_df.iterrows():
            member_id = _clean_cell(row.get("member_id"))
            member_name = normalize_search_text(row.get("member_name", ""))
            for column in cluster_columns:
                cluster = _clean_cell(column)
                score = _float_cell(row.get(column), math.nan)
                if not cluster or math.isnan(score):
                    continue
                if member_id:
                    relevance[(f"id:{member_id}", cluster)] = score
                if member_name:
                    relevance[(f"name:{member_name}", cluster)] = score
    return relevance


def _enrich_candidate_relevance_from_matrix(df: pd.DataFrame, paths: list[Path]) -> pd.DataFrame:
    if df.empty:
        return df
    relevance = _load_static_member_cluster_matrix(paths)
    if not relevance:
        return df
    df = df.copy()
    filled = 0
    for idx, row in df.iterrows():
        cluster = _clean_cell(row.get("concrete_cluster"))
        if not cluster:
            continue
        for side in ("a", "b"):
            rel_col = f"relevance_{side}"
            current = _float_cell(row.get(rel_col), math.nan)
            if not math.isnan(current) and current > 0:
                continue
            member_id = _clean_cell(row.get(f"member_{side}_id"))
            member_name = normalize_search_text(row.get(f"member_{side}", ""))
            matrix_value = math.nan
            if member_id:
                matrix_value = relevance.get((f"id:{member_id}", cluster), math.nan)
            if (math.isnan(matrix_value) or matrix_value <= 0) and member_name:
                matrix_value = relevance.get((f"name:{member_name}", cluster), math.nan)
            if not math.isnan(matrix_value) and matrix_value > 0:
                df.at[idx, rel_col] = matrix_value
                filled += 1
    if filled:
        print(f"Static Excel matrix: filled {filled} missing relevance values from {STATIC_MEMBER_CLUSTER_MATRIX_SHEET}")
    return df


def _load_static_candidate_dataframe(path: Path) -> tuple[pd.DataFrame, list[Path], str]:
    """Load the Excel-defined fixed pattern universe P.

    New nested-vector workbooks expose Pair_Vector as the audit universe and
    SCIP_Optimization_Pool as the bounded solver frontier. Older workbooks can
    still expose All_Pairings_5253 or Candidate_Patterns as fallbacks.
    """
    for sheet_name, mode in [
        (STATIC_SCIP_INPUT_SHEET, "nested_scip_optimization_pool"),
        (STATIC_FRONTIER_SHEET, "nested_scip_frontier"),
        (STATIC_PAIR_VECTOR_SHEET, "nested_pair_vector_full_universe"),
    ]:
        nested = _read_excel_sheet_optional(path, sheet_name)
        if nested.empty or not {"member_a", "member_b"}.issubset(nested.columns):
            continue
        df = _normalize_nested_vector_space_dataframe(path, nested, sheet_name)
        used_paths = [path]
        if not df.empty:
            subset_cols = [c for c in ["member_a", "member_b", "concrete_cluster", "opportunity_key"] if c in df.columns]
            if subset_cols and sheet_name != STATIC_SCIP_INPUT_SHEET:
                df = df.drop_duplicates(subset=subset_cols, keep="last")
            df = _unique_static_candidate_ids(df)
            if STATIC_CANDIDATE_LIMIT > 0:
                df = df.head(STATIC_CANDIDATE_LIMIT)
        return df, used_paths, mode

    full_universe = _read_excel_sheet_optional(path, STATIC_FULL_PAIRING_SHEET) if USE_STATIC_FULL_PAIRING_UNIVERSE else pd.DataFrame()
    if not full_universe.empty and {"member_a", "member_b"}.issubset(full_universe.columns):
        df = _normalize_full_pairing_universe_dataframe(full_universe)
        df = _candidate_rows_without_removed_members(df)
        used_paths = [path]
        mode = "full_pairing_universe"
        if not df.empty:
            df = _enrich_candidate_relevance_from_matrix(df, used_paths)
            subset_cols = [c for c in ["member_a", "member_b", "concrete_cluster", "opportunity_key"] if c in df.columns]
            if subset_cols:
                df = df.drop_duplicates(subset=subset_cols, keep="last")
            df = _unique_static_candidate_ids(df)
            if STATIC_CANDIDATE_LIMIT > 0:
                df = df.head(STATIC_CANDIDATE_LIMIT)
        return df, used_paths, mode

    preferred = _read_excel_sheet_optional(path, STATIC_CANDIDATE_SHEET)
    used_paths: list[Path] = []
    mode = "preferred_workbook"
    if not preferred.empty:
        used_paths.append(path)
    df = preferred

    preferred_looks_like_patch = len(preferred) > 0 and len(preferred) < STATIC_PATCH_MIN_CANDIDATES
    base_available = STATIC_CANDIDATE_BASE_PATH.exists() and STATIC_CANDIDATE_BASE_PATH.resolve() != path.resolve()
    if preferred_looks_like_patch and base_available:
        base = _read_excel_sheet_optional(STATIC_CANDIDATE_BASE_PATH, STATIC_CANDIDATE_SHEET)
        if not base.empty:
            base = _candidate_rows_without_removed_members(base)
            df = pd.concat([base, preferred], ignore_index=True, sort=False)
            used_paths = [STATIC_CANDIDATE_BASE_PATH, path]
            mode = "base_plus_update_patch"
    elif df.empty and base_available:
        base = _read_excel_sheet_optional(STATIC_CANDIDATE_BASE_PATH, STATIC_CANDIDATE_SHEET)
        if not base.empty:
            df = base
            used_paths = [STATIC_CANDIDATE_BASE_PATH]
            mode = "base_fallback"

    df = _candidate_rows_without_removed_members(df)
    if not df.empty:
        df = _merge_candidate_role_balance(path, df)
        df = _enrich_candidate_relevance_from_matrix(df, used_paths or [path])
        subset_cols = [c for c in ["member_a", "member_b", "concrete_cluster", "opportunity_key"] if c in df.columns]
        if subset_cols:
            df = df.drop_duplicates(subset=subset_cols, keep="last")
        df = _unique_static_candidate_ids(df)
        if STATIC_CANDIDATE_LIMIT > 0:
            df = df.head(STATIC_CANDIDATE_LIMIT)
    return df, used_paths, mode


def _numeric_rank_value(df: pd.DataFrame, *columns: str) -> pd.Series:
    for column in columns:
        if column in df.columns:
            return pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return pd.Series(0.0, index=df.index)


def _stable_row_hash(value: Any) -> int:
    digest = hashlib.sha1(str(value or "").encode("utf-8", errors="ignore")).hexdigest()[:12]
    return int(digest, 16)


def _frontier_live_signal_scores(frontier: pd.DataFrame, events_df: pd.DataFrame) -> pd.Series:
    if frontier.empty or events_df.empty or "mitglied" not in events_df.columns:
        return pd.Series(0.0, index=frontier.index)

    event_scores: dict[str, float] = defaultdict(float)
    for _, event in events_df.iterrows():
        member_key = normalize_search_text(event.get("mitglied", ""))
        if not member_key:
            continue
        event_scores[member_key] += max(
            _float_cell(event.get("decision_score"), 0.0),
            _float_cell(event.get("score"), 0.0),
            1.0,
        )

    scores: list[float] = []
    for _, row in frontier.iterrows():
        a_score = event_scores.get(normalize_search_text(row.get("member_a", "")), 0.0)
        b_score = event_scores.get(normalize_search_text(row.get("member_b", "")), 0.0)
        balanced_bonus = 12.0 if a_score > 0 and b_score > 0 else 0.0
        one_sided_bonus = 4.0 if (a_score > 0 or b_score > 0) and not balanced_bonus else 0.0
        scores.append(min(40.0, (a_score + b_score) / 12.0 + balanced_bonus + one_sided_bonus))
    return pd.Series(scores, index=frontier.index)


def _feedback_neighbor_scores(frontier: pd.DataFrame) -> pd.Series:
    if frontier.empty:
        return pd.Series(dtype=float)
    try:
        records = load_feedback_records()
    except Exception:
        records = []
    if not records:
        return pd.Series(0.0, index=frontier.index)

    pair_weights: Counter[str] = Counter()
    actor_weights: Counter[str] = Counter()
    opportunity_weights: Counter[str] = Counter()
    label_weight = {
        "useful": 4.0,
        "interesting_but_weak": 2.0,
        "good_topic_wrong_actors": 1.5,
        "wrong_actors": 1.5,
        "wrong_connection": 1.0,
        "not_relevant": 1.0,
    }
    for record in records:
        label = feedback_record_label(record)
        weight = label_weight.get(label, 1.0)
        pair = normalize_search_text(record.get("pair") or record.get("members") or record.get("member_pair_key", ""))
        if pair:
            pair_weights[pair] += weight
        opportunity = normalize_search_text(record.get("opportunity_id") or record.get("opportunity_key", ""))
        if opportunity:
            opportunity_weights[opportunity] += weight
        actors = record.get("actors")
        if not isinstance(actors, list):
            actors = re.split(r"\s*[|x×,]\s*", str(record.get("pair") or record.get("members") or ""))
        for actor in actors:
            actor_key = normalize_search_text(actor)
            if actor_key:
                actor_weights[actor_key] += max(0.5, weight * 0.35)

    scores: list[float] = []
    for _, row in frontier.iterrows():
        pair_key = normalize_search_text(row.get("member_pair_key", ""))
        opportunity_key = normalize_search_text(row.get("best_opportunity_key", ""))
        score = pair_weights.get(pair_key, 0.0) * 8.0
        score += opportunity_weights.get(opportunity_key, 0.0) * 5.0
        score += actor_weights.get(normalize_search_text(row.get("member_a", "")), 0.0)
        score += actor_weights.get(normalize_search_text(row.get("member_b", "")), 0.0)
        scores.append(score)
    return pd.Series(scores, index=frontier.index)


def _merge_pair_vector_metadata(frontier: pd.DataFrame, pair_vector: pd.DataFrame) -> pd.DataFrame:
    if frontier.empty or pair_vector.empty or "pair_id" not in frontier.columns or "pair_id" not in pair_vector.columns:
        return frontier
    meta_cols = [
        "pair_id", "subrole_a", "subrole_b", "role_a", "role_b", "final_scip_score",
        "base_structural_score", "evidence_gate_multiplier", "geo_proximity_score",
        "shared_cluster_count", "weighted_cluster_overlap", "cluster_similarity_score",
        "old_candidate_max_score", "old_candidate_rows", "filter_reason",
    ]
    meta_cols = [column for column in meta_cols if column in pair_vector.columns]
    if len(meta_cols) <= 1:
        return frontier
    meta = pair_vector[meta_cols].drop_duplicates(subset=["pair_id"], keep="last")
    return frontier.merge(meta, on="pair_id", how="left", suffixes=("", "_pair"))


def _append_unique_pool_rows(target: list[int], candidates: list[int], seen: set[int], limit: int) -> None:
    for idx in candidates:
        if idx in seen:
            continue
        target.append(idx)
        seen.add(idx)
        if len(target) >= limit:
            break


def refresh_static_scip_optimization_pool(workbook_path: Path, events_df: pd.DataFrame) -> dict[str, Any]:
    """Rebuild SCIP_Optimization_Pool from the frontier before every optimizer run."""
    if not REFRESH_STATIC_SCIP_POOL or not workbook_path.exists():
        return {"status": "skipped", "reason": "disabled_or_missing_workbook"}

    frontier_raw = _read_excel_sheet_optional(workbook_path, STATIC_FRONTIER_SHEET)
    if frontier_raw.empty or "member_a" not in frontier_raw.columns or "member_b" not in frontier_raw.columns:
        return {"status": "skipped", "reason": "missing_frontier"}

    pair_vector = _read_excel_sheet_optional(workbook_path, STATIC_PAIR_VECTOR_SHEET)
    frontier = _merge_pair_vector_metadata(frontier_raw.copy(), pair_vector)
    frontier = frontier.copy()
    frontier["_frontier_score"] = _numeric_rank_value(frontier, "frontier_score", "final_scip_score", "card_score")
    frontier["_evidence_strength"] = _numeric_rank_value(frontier, "evidence_strength", "legacy_score", "old_candidate_max_score")
    frontier["_topic_fit"] = _numeric_rank_value(frontier, "topic_fit", "cluster_similarity_score", "weighted_cluster_overlap")
    frontier["_role_fit"] = _numeric_rank_value(frontier, "role_fit", "role_complementarity_score")
    frontier["_subrole_fit"] = _numeric_rank_value(frontier, "subrole_fit", "subrole_complementarity_score")
    frontier["_geo_fit"] = _numeric_rank_value(frontier, "geo_fit", "geo_proximity_score")
    frontier["_protected"] = _numeric_rank_value(frontier, "protected_candidate")
    frontier["_weak"] = _numeric_rank_value(frontier, "weak_evidence_flag")
    frontier["_low"] = _numeric_rank_value(frontier, "low_score_flag")
    frontier["_dominated"] = _numeric_rank_value(frontier, "dominated_low_value")
    frontier["_live_score"] = _frontier_live_signal_scores(frontier, events_df)
    frontier["_feedback_neighbor_score"] = _feedback_neighbor_scores(frontier)
    frontier["_stable_hash"] = frontier.apply(
        lambda row: _stable_row_hash(row.get("frontier_id") or row.get("pair_id") or row.get("member_pair_key")),
        axis=1,
    )
    frontier["_refresh_score"] = (
        frontier["_frontier_score"]
        + frontier["_live_score"]
        + frontier["_feedback_neighbor_score"] * 0.8
        + frontier["_protected"] * 8.0
        - frontier["_weak"] * 4.0
        - frontier["_low"] * 5.0
        - frontier["_dominated"] * 15.0
    )

    eligible = frontier[frontier["_dominated"] <= 0].copy()
    if eligible.empty:
        eligible = frontier.copy()

    selected_indices: list[int] = []
    seen: set[int] = set()
    high_candidates = eligible.sort_values(
        ["_refresh_score", "_frontier_score", "_evidence_strength", "_stable_hash"],
        ascending=[False, False, False, True],
    ).index.tolist()
    _append_unique_pool_rows(selected_indices, high_candidates, seen, min(SCIP_POOL_HIGH_SCORE_COUNT, SCIP_POOL_TARGET_COUNT))

    feedback_candidates = eligible[eligible["_feedback_neighbor_score"] > 0].sort_values(
        ["_feedback_neighbor_score", "_refresh_score", "_stable_hash"],
        ascending=[False, False, True],
    ).index.tolist()
    _append_unique_pool_rows(
        selected_indices,
        feedback_candidates,
        seen,
        min(SCIP_POOL_TARGET_COUNT, len(selected_indices) + SCIP_POOL_FEEDBACK_NEIGHBOR_COUNT),
    )

    exploration_target = min(SCIP_POOL_TARGET_COUNT, len(selected_indices) + SCIP_POOL_EXPLORATION_COUNT)
    exploration_candidates: list[int] = []
    subrole_cols = [column for column in ["subrole_a", "subrole_b", "role_a", "role_b"] if column in eligible.columns]
    if subrole_cols:
        exploded: list[tuple[int, int, float, int, int]] = []
        subrole_counts: Counter[str] = Counter()
        for idx in selected_indices:
            row = eligible.loc[idx] if idx in eligible.index else None
            if row is None:
                continue
            for column in subrole_cols[:2]:
                value = _clean_cell(row.get(column))
                if value:
                    subrole_counts[value] += 1
        for idx, row in eligible.iterrows():
            if idx in seen:
                continue
            values = [_clean_cell(row.get(column)) for column in subrole_cols if _clean_cell(row.get(column))]
            rarity = min((subrole_counts.get(value, 0) for value in values), default=0)
            exploded.append((rarity, int(row["_weak"] + row["_low"]), -float(row["_refresh_score"]), int(row["_stable_hash"]), idx))
        exploration_candidates = [item[-1] for item in sorted(exploded)]
    else:
        exploration_candidates = eligible.sort_values(["_stable_hash"], ascending=True).index.tolist()
    _append_unique_pool_rows(selected_indices, exploration_candidates, seen, exploration_target)

    if len(selected_indices) < SCIP_POOL_TARGET_COUNT:
        fill_candidates = frontier.sort_values(
            ["_refresh_score", "_frontier_score", "_stable_hash"],
            ascending=[False, False, True],
        ).index.tolist()
        _append_unique_pool_rows(selected_indices, fill_candidates, seen, SCIP_POOL_TARGET_COUNT)

    pool = frontier.loc[selected_indices].copy()
    output_cols = [column for column in frontier_raw.columns if column in pool.columns]
    pool = pool[output_cols].copy()
    pool["pool_rank"] = range(1, len(pool) + 1)
    pool["pool_lane"] = [
        "high_score" if idx < SCIP_POOL_HIGH_SCORE_COUNT
        else "feedback_neighbor" if idx < SCIP_POOL_HIGH_SCORE_COUNT + SCIP_POOL_FEEDBACK_NEIGHBOR_COUNT
        else "exploration"
        for idx in range(len(pool))
    ]
    pool["pool_refresh_score"] = frontier.loc[selected_indices, "_refresh_score"].round(4).to_list()
    pool["pool_live_signal_score"] = frontier.loc[selected_indices, "_live_score"].round(4).to_list()
    pool["pool_feedback_neighbor_score"] = frontier.loc[selected_indices, "_feedback_neighbor_score"].round(4).to_list()

    with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        pool.to_excel(writer, sheet_name=STATIC_SCIP_INPUT_SHEET, index=False)

    lane_counts = dict(Counter(pool["pool_lane"]))
    print(
        "Static Excel SCIP pool refreshed: "
        f"{len(pool)} rows ({lane_counts}) from {STATIC_FRONTIER_SHEET}; "
        f"full_pair_universe={len(pair_vector) if not pair_vector.empty else 'unknown'}"
    )
    return {
        "status": "refreshed",
        "rows": int(len(pool)),
        "lane_counts": lane_counts,
        "frontier_rows": int(len(frontier_raw)),
        "pair_vector_rows": int(len(pair_vector)) if not pair_vector.empty else None,
    }


def _load_static_constraint_limits_from_workbooks(paths: list[Path]) -> tuple[dict[str, int], dict[str, int]]:
    member_limits: dict[str, int] = {}
    cluster_limits: dict[str, int] = {}
    for workbook in paths:
        members_df = _read_excel_sheet_optional(workbook, STATIC_MEMBER_CONSTRAINT_SHEET)
        if not members_df.empty:
            for _, row in members_df.iterrows():
                member = normalize_search_text(row.get("member_name", row.get("mitglied", "")))
                if not member or member in STATIC_REMOVED_MEMBERS:
                    continue
                try:
                    member_limits[member] = int(row.get("limit", MAX_MEETINGS_PER_MEMBER) or MAX_MEETINGS_PER_MEMBER)
                except Exception:
                    member_limits[member] = MAX_MEETINGS_PER_MEMBER
        clusters_df = _read_excel_sheet_optional(workbook, STATIC_CLUSTER_CONSTRAINT_SHEET)
        if not clusters_df.empty:
            for _, row in clusters_df.iterrows():
                cluster = _clean_cell(row.get("concrete_cluster"))
                if not cluster:
                    continue
                try:
                    cluster_limits[cluster] = int(row.get("limit", MAX_SELECTED_PER_CLUSTER) or MAX_SELECTED_PER_CLUSTER)
                except Exception:
                    cluster_limits[cluster] = MAX_SELECTED_PER_CLUSTER
    return member_limits, cluster_limits


def _load_role_quota_from_workbooks(paths: list[Path]) -> tuple[dict[str, float], dict[str, float], dict[str, tuple[str, float]]]:
    role_limits: dict[str, float] = {}
    role_inverse: dict[str, float] = {}
    member_weights: dict[str, tuple[str, float]] = {}
    for workbook in paths:
        roles_df = _read_excel_table_optional(
            workbook,
            "Role_Weights",
            {"actor_role", "recommended_max_role_mass"},
        )
        if not roles_df.empty:
            for _, row in roles_df.iterrows():
                role = _clean_cell(row.get("actor_role"))
                if not role:
                    continue
                limit = _float_cell(row.get("recommended_max_role_mass"), math.nan)
                if not math.isnan(limit) and limit > 0:
                    role_limits[role] = limit
                inverse = _float_cell(row.get("member_inverse_weight = 1/n_role"), math.nan)
                if math.isnan(inverse):
                    inverse = _float_cell(row.get("member_inverse_weight"), math.nan)
                if not math.isnan(inverse) and inverse > 0:
                    role_inverse[role] = inverse

        members_df = _read_excel_table_optional(
            workbook,
            "Member_Role_Weights",
            {"member_name", "actor_role", "member_inverse_weight"},
        )
        if not members_df.empty:
            for _, row in members_df.iterrows():
                member = _clean_cell(row.get("member_name"))
                role = _clean_cell(row.get("actor_role"))
                weight = _float_cell(row.get("member_inverse_weight"), math.nan)
                if member and role and not math.isnan(weight) and weight > 0:
                    member_weights[normalize_search_text(member)] = (role, weight)
    return role_limits, role_inverse, member_weights


def _load_subrole_quota_from_workbooks(paths: list[Path]) -> tuple[dict[str, float], dict[str, float], dict[str, tuple[str, float]]]:
    subrole_limits: dict[str, float] = {}
    subrole_inverse: dict[str, float] = {}
    member_weights: dict[str, tuple[str, float]] = {}
    for workbook in paths:
        subroles_df = _read_excel_table_optional(
            workbook,
            "Subrole_Weights",
            {"actor_subrole", "recommended_max_subrole_mass"},
        )
        if not subroles_df.empty:
            for _, row in subroles_df.iterrows():
                subrole = _clean_cell(row.get("actor_subrole"))
                if not subrole:
                    continue
                limit = _float_cell(row.get("recommended_max_subrole_mass"), math.nan)
                if not math.isnan(limit) and limit > 0:
                    subrole_limits[subrole] = limit
                inverse = _float_cell(row.get("member_inverse_weight = 1/n_subrole"), math.nan)
                if math.isnan(inverse):
                    inverse = _float_cell(row.get("subrole_weight"), math.nan)
                if math.isnan(inverse):
                    inverse = _float_cell(row.get("member_inverse_weight"), math.nan)
                if not math.isnan(inverse) and inverse > 0:
                    subrole_inverse[subrole] = inverse

        members_df = _read_excel_table_optional(
            workbook,
            "Member_Subrole_Weights",
            {"member_name", "actor_subrole", "subrole_weight"},
        )
        if members_df.empty:
            members_df = _read_excel_table_optional(
                workbook,
                "Member_Subrole_Weights",
                {"member_name", "actor_subrole", "member_inverse_weight"},
            )
        if not members_df.empty:
            for _, row in members_df.iterrows():
                member = _clean_cell(row.get("member_name"))
                subrole = _clean_cell(row.get("actor_subrole"))
                weight = _float_cell(row.get("subrole_weight"), math.nan)
                if math.isnan(weight):
                    weight = _float_cell(row.get("member_inverse_weight"), math.nan)
                if member and subrole and not math.isnan(weight) and weight > 0:
                    member_weights[normalize_search_text(member)] = (subrole, weight)
    return subrole_limits, subrole_inverse, member_weights


def _load_portfolio_constraints_from_workbooks(paths: list[Path]) -> tuple[int, int, int, int, int, float, int]:
    max_academic_matches = STATIC_MAX_ACADEMIC_MATCHES
    max_academic_actor_slots = STATIC_MAX_ACADEMIC_ACTOR_SLOTS
    min_non_academic_pairs = STATIC_MIN_NON_ACADEMIC_PAIRS
    max_weak_evidence_matches = STATIC_MAX_WEAK_EVIDENCE_MATCHES
    max_low_score_matches = STATIC_MAX_LOW_SCORE_MATCHES
    min_geo_proximity_sum = STATIC_MIN_GEO_PROXIMITY_SUM
    max_academic_academic_pairs = STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS
    for workbook in paths:
        df = _read_excel_table_optional(
            workbook,
            "SCIP_Portfolio_Constraints",
            {"constraint_name", "limit_value"},
        )
        if df.empty:
            df = _read_excel_table_optional(
                workbook,
                "SCIP_Portfolio_Constraints",
                {"key", "value"},
            )
        if df.empty:
            continue
        for _, row in df.iterrows():
            key = _clean_cell(row.get("constraint_name") or row.get("key")).lower()
            if not key:
                continue
            enabled = _float_cell(row.get("enabled"), 1.0)
            if not math.isnan(enabled) and enabled <= 0:
                continue
            value = _float_cell(row.get("limit_value"), math.nan)
            if math.isnan(value):
                value = _float_cell(row.get("value"), math.nan)
            if math.isnan(value):
                continue
            if key == "max_academic_matches":
                max_academic_matches = max(0, int(value))
            elif key == "max_academic_actor_slots":
                max_academic_actor_slots = max(0, int(value))
            elif key == "min_non_academic_pairs":
                min_non_academic_pairs = max(0, int(value))
            elif key == "max_weak_evidence_matches":
                max_weak_evidence_matches = max(0, int(value))
            elif key == "max_low_score_matches":
                max_low_score_matches = max(0, int(value))
            elif key == "min_geo_proximity_sum":
                min_geo_proximity_sum = max(0.0, float(value))
            elif key == "max_academic_academic_pairs":
                max_academic_academic_pairs = max(0, int(value))
    return (
        max_academic_matches,
        max_academic_actor_slots,
        min_non_academic_pairs,
        max_weak_evidence_matches,
        max_low_score_matches,
        min_geo_proximity_sum,
        max_academic_academic_pairs,
    )

def _member_event_rows(events_df: pd.DataFrame, member: str) -> pd.DataFrame:
    if events_df.empty or "mitglied" not in events_df.columns:
        return events_df.iloc[0:0]
    member_key = normalize_search_text(member)
    if not member_key:
        return events_df.iloc[0:0]
    keys = events_df["mitglied"].map(normalize_search_text)
    return events_df[keys == member_key]


def _event_text(row: pd.Series) -> str:
    parts = []
    for column in [
        "titel", "title", "snippet", "hauptsektion", "themenfelder", "opportunity_keywords",
        "one_sentence_context", "why_it_matters", "mechanism", "implementation_angle",
        "scip_reasoning", "briefing_sentence",
    ]:
        if column in row.index:
            parts.append(str(row.get(column, "") or ""))
    return normalize_search_text(" ".join(parts))


def _event_title(row: pd.Series) -> str:
    return _clean_cell(row.get("titel", "") or row.get("title", ""))


def _static_candidate_live_evidence(
    member_a: str,
    member_b: str,
    concrete_cluster: str,
    themenfeld: str,
    events_df: pd.DataFrame,
    cluster_keywords: dict[str, set[str]],
) -> dict[str, Any]:
    keywords = set(cluster_keywords.get(concrete_cluster, set()))
    if concrete_cluster:
        keywords.add(normalize_search_text(concrete_cluster).replace("_", " "))
    if themenfeld:
        keywords.add(normalize_search_text(themenfeld))
    keywords = {kw for kw in keywords if len(kw) >= 3}

    rows: list[pd.Series] = []
    per_member_counts = {member_a: 0, member_b: 0}
    for member in [member_a, member_b]:
        member_rows = _member_event_rows(events_df, member)
        scored_rows: list[tuple[float, pd.Series]] = []
        for _, event in member_rows.iterrows():
            text = _event_text(event)
            keyword_hits = sum(1 for kw in keywords if kw and kw in text)
            score = _float_cell(event.get("decision_score", event.get("score", 0)), 0.0)
            if keyword_hits or not keywords:
                scored_rows.append((keyword_hits * 10.0 + score, event))
        scored_rows.sort(key=lambda item: item[0], reverse=True)
        for _, event in scored_rows[:3]:
            rows.append(event)
            per_member_counts[member] += 1

    titles: list[str] = []
    urls: list[str] = []
    live_scores: list[float] = []
    for event in rows[:6]:
        title = _event_title(event)
        event_member = _clean_cell(event.get("mitglied"))
        if title:
            titles.append(f"{event_member}: {title}" if event_member else title)
        url = _clean_cell(event.get("url", "") or event.get("seite", ""))
        if url:
            urls.append(url)
        live_scores.append(_float_cell(event.get("decision_score", event.get("score", 0)), 0.0))

    event_count = len(rows)
    both_members = int(per_member_counts.get(member_a, 0) > 0) + int(per_member_counts.get(member_b, 0) > 0)
    avg_decision = sum(live_scores) / len(live_scores) if live_scores else 0.0
    boost = min(18.0, event_count * 2.0 + both_members * 2.5 + max(0.0, avg_decision - 30.0) / 10.0)
    if event_count == 0:
        status = "no_live_evidence"
    elif both_members >= 2:
        status = "balanced_live_evidence"
    else:
        status = "one_sided_live_evidence"
    return {
        "event_count": event_count,
        "member_a_count": per_member_counts.get(member_a, 0),
        "member_b_count": per_member_counts.get(member_b, 0),
        "avg_decision_score": round(avg_decision, 3),
        "boost": round(boost, 3),
        "titles": tuple(titles),
        "urls": tuple(urls),
        "status": status,
    }


def load_static_candidate_patterns_from_excel(path: Path, events_df: pd.DataFrame) -> list[MeetingPattern]:
    global STATIC_MEMBER_LIMITS, STATIC_CLUSTER_LIMITS, STATIC_ROLE_MASS_LIMITS
    global STATIC_ROLE_INVERSE_WEIGHTS, STATIC_MEMBER_ROLE_WEIGHTS
    global STATIC_SUBROLE_MASS_LIMITS, STATIC_SUBROLE_INVERSE_WEIGHTS, STATIC_MEMBER_SUBROLE_WEIGHTS
    global STATIC_MAX_ACADEMIC_MATCHES, STATIC_MAX_ACADEMIC_ACTOR_SLOTS
    global STATIC_MIN_NON_ACADEMIC_PAIRS, STATIC_MAX_WEAK_EVIDENCE_MATCHES
    global STATIC_MAX_LOW_SCORE_MATCHES, STATIC_MIN_GEO_PROXIMITY_SUM
    global STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS
    df, used_paths, mode = _load_static_candidate_dataframe(path)
    if df.empty:
        raise FileNotFoundError(f"No static candidate patterns could be loaded from {path}")

    STATIC_MEMBER_LIMITS, STATIC_CLUSTER_LIMITS = _load_static_constraint_limits_from_workbooks(used_paths or [path])
    STATIC_ROLE_MASS_LIMITS, STATIC_ROLE_INVERSE_WEIGHTS, STATIC_MEMBER_ROLE_WEIGHTS = _load_role_quota_from_workbooks(used_paths or [path])
    STATIC_SUBROLE_MASS_LIMITS, STATIC_SUBROLE_INVERSE_WEIGHTS, STATIC_MEMBER_SUBROLE_WEIGHTS = _load_subrole_quota_from_workbooks(used_paths or [path])
    (
        STATIC_MAX_ACADEMIC_MATCHES,
        STATIC_MAX_ACADEMIC_ACTOR_SLOTS,
        STATIC_MIN_NON_ACADEMIC_PAIRS,
        STATIC_MAX_WEAK_EVIDENCE_MATCHES,
        STATIC_MAX_LOW_SCORE_MATCHES,
        STATIC_MIN_GEO_PROXIMITY_SUM,
        STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS,
    ) = _load_portfolio_constraints_from_workbooks(used_paths or [path])
    cluster_keywords: dict[str, set[str]] = {}
    for workbook in used_paths or [path]:
        cluster_keywords.update(_static_cluster_keyword_map(workbook))

    patterns: list[MeetingPattern] = []
    for fallback_id, (_, row) in enumerate(df.iterrows(), start=1):
        row_dict = row.to_dict()
        candidate_id = _clean_cell(row_dict.get("candidate_id")) or f"C{fallback_id:04d}"
        scip_variable = _clean_cell(row_dict.get("scip_variable")) or f"x_{fallback_id:04d}"
        member_a = _clean_cell(row_dict.get("member_a"))
        member_b = _clean_cell(row_dict.get("member_b"))
        if not member_a or not member_b:
            continue
        if (
            not mode.startswith("nested_")
            and (normalize_search_text(member_a) in STATIC_REMOVED_MEMBERS or normalize_search_text(member_b) in STATIC_REMOVED_MEMBERS)
        ):
            continue

        themenfeld = _clean_cell(row_dict.get("themenfeld"))
        concrete_cluster = _clean_cell(row_dict.get("concrete_cluster"))
        cluster_label = _clean_cell(row_dict.get("cluster_label")) or display_topic(concrete_cluster)
        static_problem = _clean_cell(row_dict.get("static_cluster_problem"))
        reasoning = _clean_cell(row_dict.get("reasoning"))
        role_a = _clean_cell(row_dict.get("role_a"))
        role_b = _clean_cell(row_dict.get("role_b"))
        static_score = _static_support_score_from_row(row_dict)
        opportunity_key = _clean_cell(row_dict.get("opportunity_key"))
        member_pair_key = _clean_cell(row_dict.get("member_pair_key"))
        candidate_status = _clean_cell(row_dict.get("candidate_status"))
        filter_reason = _clean_cell(row_dict.get("filter_reason"))
        geo_score_reason = _clean_cell(row_dict.get("geo_score_reason"))
        geo_proximity_score = _float_cell(row_dict.get("geo_proximity_score"), 0.0)
        geo_proximity_norm = _float_cell(row_dict.get("geo_proximity_norm"), math.nan)
        if math.isnan(geo_proximity_norm):
            geo_proximity_norm = geo_proximity_score / 100.0 if geo_proximity_score > 1.0 else geo_proximity_score
        distance_km = _float_cell(row_dict.get("distance_km"), 0.0)
        same_geo_region = bool(_float_cell(row_dict.get("same_geo_region"), 0.0) > 0)
        weak_evidence_flag = int(_float_cell(row_dict.get("weak_evidence_flag"), 0.0) > 0)
        low_score_flag = int(_float_cell(row_dict.get("low_score_flag"), 0.0) > 0)
        academic_pair_flag = int(_float_cell(row_dict.get("academic_pair_flag"), _float_cell(row_dict.get("includes_academic"), 0.0)) > 0)
        academic_academic_flag = int(_float_cell(row_dict.get("academic_academic_flag"), 0.0) > 0)
        non_academic_pair_flag = int(_float_cell(row_dict.get("non_academic_pair_flag"), 0.0) > 0)
        legacy_candidate_id = _clean_cell(row_dict.get("legacy_candidate_id") or row_dict.get("old_candidate_id"))
        legacy_opportunity_key = _clean_cell(row_dict.get("legacy_opportunity_key"))
        old_candidate_max_score = _float_cell(row_dict.get("old_candidate_max_score"), 0.0)
        old_candidate_rows = int(_float_cell(row_dict.get("old_candidate_rows"), 0.0))
        live = _static_candidate_live_evidence(member_a, member_b, concrete_cluster, themenfeld, events_df, cluster_keywords)
        score = max(0.0, static_score + float(live["boost"]))
        members = tuple(sorted([member_a, member_b]))
        why = (
            f"Static reason: {reasoning}" if reasoning else
            f"Static reason: {member_a} and {member_b} are Excel-defined for {cluster_label}."
        )
        editorial = (
            f"Excel mode: {mode}. Live evidence: {live['status']} with {live['event_count']} current signals. "
            f"Static cluster problem: {static_problem}"
        ).strip()
        agenda = (
            f"Validate the Excel-defined opportunity '{cluster_label}' against current evidence, roles and a concrete RLP application anchor."
        )
        patterns.append(MeetingPattern(
            pattern_id=fallback_id if mode.startswith("nested_") else _int_from_candidate_id(candidate_id, fallback_id),
            members=members,
            score=score,
            shared_topics=tuple(t for t in [themenfeld, concrete_cluster] if t),
            bridge_topics=tuple(t for t in [_clean_cell(row_dict.get("complementary_topic_pairs"))] if t),
            concrete_clusters=tuple(t for t in [concrete_cluster] if t),
            relevant_names=tuple(t for t in [role_a, role_b] if t),
            reason=reasoning or f"{cluster_label}: {member_a} + {member_b}",
            convening_theme=cluster_label,
            why_these_members=why,
            editorial_justification=editorial,
            recent_signals=tuple(live["titles"]),
            suggested_format="Kuratiertes Sondierungsgespraech" if live["event_count"] else "Feedback-/Watchlist-Kandidat",
            possible_agenda=agenda,
            candidate_id=candidate_id,
            scip_variable=scip_variable,
            candidate_source="static_excel",
            opportunity_key=opportunity_key,
            member_pair_key=member_pair_key,
            static_support_score=static_score,
            live_signal_boost=float(live["boost"]),
            live_event_count=int(live["event_count"]),
            live_signal_titles=tuple(live["titles"]),
            live_signal_urls=tuple(live["urls"]),
            live_evidence_status=str(live["status"]),
            geo_proximity_score=geo_proximity_score,
            geo_proximity_norm=geo_proximity_norm,
            distance_km=distance_km,
            same_geo_region=same_geo_region,
            geo_score_reason=geo_score_reason,
            candidate_status=candidate_status,
            filter_reason=filter_reason,
            weak_evidence_flag=weak_evidence_flag,
            low_score_flag=low_score_flag,
            academic_pair_flag=academic_pair_flag,
            academic_academic_flag=academic_academic_flag,
            non_academic_pair_flag=non_academic_pair_flag,
            legacy_candidate_id=legacy_candidate_id,
            legacy_opportunity_key=legacy_opportunity_key,
            old_candidate_max_score=old_candidate_max_score,
            old_candidate_rows=old_candidate_rows,
        ))
    patterns.sort(key=lambda p: (-p.score, p.candidate_id or str(p.pattern_id)))
    print(
        "Static Excel loader: "
        f"mode={mode}, rows={len(df)}, patterns={len(patterns)}, "
        f"workbooks={', '.join(str(p) for p in used_paths or [path])}"
    )
    print(
        "Static Excel constraints: "
        f"member_limits={len(STATIC_MEMBER_LIMITS)}, cluster_limits={len(STATIC_CLUSTER_LIMITS)}, "
        f"role_quota_limits={len(STATIC_ROLE_MASS_LIMITS)}, "
        f"subrole_quota_limits={len(STATIC_SUBROLE_MASS_LIMITS)}, "
        f"academic_caps={STATIC_MAX_ACADEMIC_MATCHES}/{STATIC_MAX_ACADEMIC_ACTOR_SLOTS}, "
        f"non_academic_min={STATIC_MIN_NON_ACADEMIC_PAIRS}, "
        f"weak_evidence_cap={STATIC_MAX_WEAK_EVIDENCE_MATCHES}, "
        f"low_score_cap={STATIC_MAX_LOW_SCORE_MATCHES}, "
        f"geo_floor={STATIC_MIN_GEO_PROXIMITY_SUM:.2f}, "
        f"academic_academic_cap={STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS}"
    )
    return patterns

def generate_meeting_patterns(profiles: dict[str, MemberProfile]) -> list[MeetingPattern]:
    active = [p for p in profiles.values() if p.event_count > 0 and p.topics]
    patterns: list[MeetingPattern] = []
    pid = 0

    # Pairs are the base "valid patterns".
    for a, b in combinations(active, 2):
        score, shared, bridge, names = pair_score(a, b)
        if not pattern_is_editorially_credible((a, b), score, shared, bridge):
            continue
        theme, why, editorial, signals, fmt, agenda = build_convening_fields((a, b), shared, bridge)
        patterns.append(MeetingPattern(
            pattern_id=pid,
            members=tuple(sorted([a.name, b.name])),
            score=score,
            shared_topics=shared,
            bridge_topics=bridge,
            concrete_clusters=tuple(t for t in shared if cluster_info(t)),
            relevant_names=names,
            reason=reason_for_pattern((a, b), shared, bridge),
            convening_theme=theme,
            why_these_members=why,
            editorial_justification=editorial,
            recent_signals=signals,
            suggested_format=fmt,
            possible_agenda=agenda,
        ))
        pid += 1

    # Optional grouped meetings. Disabled for ranked bilateral SCIP matches.
    if not PAIR_ONLY:
        pair_lookup = {tuple(pattern.members): pattern.score for pattern in patterns if len(pattern.members) == 2}
        for size in range(3, MAX_CLUSTER_SIZE + 1):
            for group in combinations(active, size):
                member_names = tuple(sorted(p.name for p in group))
                pair_keys = [tuple(sorted((a.name, b.name))) for a, b in combinations(group, 2)]
                if not any(pair_lookup.get(k, 0) >= 10 for k in pair_keys):
                    continue
                score, shared, bridge, names = cluster_score(group)
                if not pattern_is_editorially_credible(group, score, shared, bridge):
                    continue
                theme, why, editorial, signals, fmt, agenda = build_convening_fields(group, shared, bridge)
                patterns.append(MeetingPattern(
                    pattern_id=pid,
                    members=member_names,
                    score=score,
                    shared_topics=shared,
                    bridge_topics=bridge,
                    concrete_clusters=tuple(t for t in shared if cluster_info(t)),
                    relevant_names=names,
                    reason=reason_for_pattern(group, shared, bridge),
                    convening_theme=theme,
                    why_these_members=why,
                    editorial_justification=editorial,
                    recent_signals=signals,
                    suggested_format=fmt,
                    possible_agenda=agenda,
                ))
                pid += 1

    if len(patterns) < MIN_CANDIDATE_PATTERNS:
        existing_keys = {tuple(pattern.members) for pattern in patterns}
        for a, b in combinations(active, 2):
            member_names = tuple(sorted([a.name, b.name]))
            if member_names in existing_keys:
                continue
            score, shared, bridge, names = pair_score(a, b)
            if not pattern_is_editorially_credible((a, b), score, shared, bridge, allow_backfill=True):
                continue
            theme, why, editorial, signals, fmt, agenda = build_convening_fields((a, b), shared, bridge)
            if not signals:
                continue
            patterns.append(MeetingPattern(
                pattern_id=pid,
                members=member_names,
                score=max(0.0, score - 4.0),
                shared_topics=shared,
                bridge_topics=bridge,
                concrete_clusters=tuple(t for t in shared if cluster_info(t)),
                relevant_names=names,
                reason=reason_for_pattern((a, b), shared, bridge),
                convening_theme=theme,
                why_these_members=why,
                editorial_justification=editorial,
                recent_signals=signals,
                suggested_format=fmt,
                possible_agenda=agenda,
            ))
            existing_keys.add(member_names)
            pid += 1
            if len(patterns) >= MIN_CANDIDATE_PATTERNS * 3:
                break

    patterns.sort(key=lambda p: (-final_pattern_score(p.score, p), len(p.members), p.members))
    clear_patterns = [pattern for pattern in patterns if has_clear_convening_problem(pattern)]
    if len(clear_patterns) >= MIN_CANDIDATE_PATTERNS:
        patterns = clear_patterns
    else:
        patterns = clear_patterns + [pattern for pattern in patterns if pattern not in clear_patterns]

    return [
        MeetingPattern(
            pattern_id=i,
            members=p.members,
            score=p.score,
            shared_topics=p.shared_topics,
            bridge_topics=p.bridge_topics,
            concrete_clusters=p.concrete_clusters,
            relevant_names=p.relevant_names,
            reason=p.reason,
            convening_theme=p.convening_theme,
            why_these_members=p.why_these_members,
            editorial_justification=p.editorial_justification,
            recent_signals=p.recent_signals,
            suggested_format=p.suggested_format,
            possible_agenda=p.possible_agenda,
            candidate_id=p.candidate_id,
            scip_variable=p.scip_variable,
            candidate_source=p.candidate_source,
            opportunity_key=p.opportunity_key,
            member_pair_key=p.member_pair_key,
            static_support_score=p.static_support_score,
            live_signal_boost=p.live_signal_boost,
            live_event_count=p.live_event_count,
            live_signal_titles=p.live_signal_titles,
            live_signal_urls=p.live_signal_urls,
            live_evidence_status=p.live_evidence_status,
            geo_proximity_score=p.geo_proximity_score,
            geo_proximity_norm=p.geo_proximity_norm,
            distance_km=p.distance_km,
            same_geo_region=p.same_geo_region,
            geo_score_reason=p.geo_score_reason,
            candidate_status=p.candidate_status,
            filter_reason=p.filter_reason,
            weak_evidence_flag=p.weak_evidence_flag,
            low_score_flag=p.low_score_flag,
            academic_pair_flag=p.academic_pair_flag,
            academic_academic_flag=p.academic_academic_flag,
            non_academic_pair_flag=p.non_academic_pair_flag,
            legacy_candidate_id=p.legacy_candidate_id,
            legacy_opportunity_key=p.legacy_opportunity_key,
            old_candidate_max_score=p.old_candidate_max_score,
            old_candidate_rows=p.old_candidate_rows,
        )
        for i, p in enumerate(patterns[:MAX_CANDIDATE_PATTERNS])
    ]


def has_implementation_anchor_in_pattern(pattern: MeetingPattern) -> bool:
    counts = actor_role_counts(pattern.members)
    return counts["profit_anchor"] > 0 or counts["implementation_anchor"] > 0


def filter_pair_only_patterns(patterns: list[MeetingPattern]) -> list[MeetingPattern]:
    filtered = []

    for pattern in patterns:
        if len(pattern.members) != 2:
            continue

        final_score_value = final_pattern_score(pattern.score, pattern)
        if final_score_value < MIN_FINAL_PAIR_SCORE:
            continue

        if REQUIRE_IMPLEMENTATION_ANCHOR and not has_implementation_anchor_in_pattern(pattern):
            continue

        if not has_clear_convening_problem(pattern):
            continue

        filtered.append(pattern)

    if len(filtered) < MIN_PUBLIC_CARDS:
        existing = {pattern.pattern_id for pattern in filtered}
        fallback = []
        for pattern in patterns:
            if pattern.pattern_id in existing or len(pattern.members) != 2:
                continue
            if REQUIRE_IMPLEMENTATION_ANCHOR and not has_implementation_anchor_in_pattern(pattern):
                continue
            if final_pattern_score(pattern.score, pattern) < max(18, MIN_FINAL_PAIR_SCORE - 12):
                continue
            fallback.append(pattern)
        fallback = sorted(
            fallback,
            key=lambda p: (
                not has_operational_evidence(p),
                macro_context_penalty(p),
                -final_pattern_score(p.score, p),
                p.members,
            ),
        )
        filtered.extend(fallback[:max(0, MIN_PUBLIC_CARDS - len(filtered))])

    filtered = sorted(filtered, key=lambda p: (-final_pattern_score(p.score, p), p.members))

    return filtered[:MAX_CANDIDATE_PATTERNS]


def cluster_member_relevance(member_name: str, cluster_id: str) -> float:
    relevance = CLUSTER_MEMBER_RELEVANCE.get(cluster_id, {}).get(member_name)
    if relevance is None:
        return 1.0
    return 0.90 + (float(relevance) / 5.0) * 0.25


def member_environment_score(member_name: str) -> float:
    weights = MEMBER_ENVIRONMENT_WEIGHTS.get(member_name)
    if not weights:
        return 1.0
    raw = (
        weights.get("economic_weight", 1) * 0.25
        + weights.get("regional_system_role", 1) * 0.30
        + weights.get("network_multiplier", 1) * 0.20
        + weights.get("innovation_relevance", 1) * 0.15
        + weights.get("public_interest", 1) * 0.10
    )
    return 0.90 + (raw / 5.0) * 0.25


def pattern_environment_multiplier(pattern: MeetingPattern) -> float:
    member_scores = [member_environment_score(member) for member in pattern.members]
    if not member_scores:
        return 1.0
    avg_score = sum(member_scores) / len(member_scores)
    max_score = max(member_scores)
    multiplier = min((0.75 * avg_score) + (0.25 * max_score), 1.13)
    counts = actor_role_counts(pattern.members)
    if counts["academic_support"] and not (counts["profit_anchor"] + counts["implementation_anchor"]):
        multiplier = min(multiplier, 0.96)
    return multiplier


def pattern_cluster_relevance_multiplier(pattern: MeetingPattern) -> float:
    if not pattern.concrete_clusters:
        return 1.0
    scores = []
    for topic in pattern.concrete_clusters:
        cluster_id = plain_cluster_id(topic)
        for member in pattern.members:
            scores.append(cluster_member_relevance(member, cluster_id))
    if not scores:
        return 1.0
    avg_score = sum(scores) / len(scores)
    max_score = max(scores)
    return min((0.70 * avg_score) + (0.30 * max_score), 1.14)


def can_apply_structural_bonus(pattern: MeetingPattern) -> bool:
    if not pattern.concrete_clusters:
        return False
    if not pattern.shared_topics:
        return False
    return True


def pattern_feedback_words(pattern: MeetingPattern) -> list[str]:
    text = " ".join([
        pattern.convening_theme,
        pattern.why_these_members,
        pattern.editorial_justification,
        pattern.possible_agenda,
        " ".join(pattern.shared_topics),
        " ".join(pattern.bridge_topics),
        " ".join(pattern.concrete_clusters),
        " ".join(pattern.recent_signals),
    ])
    if normalize_words is None:
        return [word.lower() for word in re.findall(r"[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\-]{2,}", text)[:40]]
    return normalize_words([], extra_text=text)


def pattern_recommendation_id(pattern: MeetingPattern) -> str:
    if stable_recommendation_id is None:
        raw = "|".join(pattern.members) + "|" + pattern.convening_theme
        import hashlib
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]
    return stable_recommendation_id(
        actors=list(pattern.members),
        topic=pattern.convening_theme,
        reason=pattern.editorial_justification,
    )


def rule_pattern_score(base_score: float, pattern: MeetingPattern) -> float:
    if not can_apply_structural_bonus(pattern):
        return max(0.0, base_score + pair_evidence_balance_bonus(pattern) - academic_support_penalty(pattern.members, final=True))
    environment_multiplier = pattern_environment_multiplier(pattern)
    cluster_multiplier = pattern_cluster_relevance_multiplier(pattern)
    combined_multiplier = (
        1.0
        + ((environment_multiplier - 1.0) * 0.40)
        + ((cluster_multiplier - 1.0) * 0.60)
    )
    if critical_action_penalty(pattern) >= 10:
        combined_multiplier = 1.0
    elif not pattern_has_application_capability_bridge(pattern):
        combined_multiplier = min(combined_multiplier, 1.02)
    elif base_score < 40:
        combined_multiplier = 1.0 + ((combined_multiplier - 1.0) * 0.25)
    elif base_score < 60:
        combined_multiplier = 1.0 + ((combined_multiplier - 1.0) * 0.60)
    scored = min(base_score * combined_multiplier, base_score + 6.0)
    scored += pair_evidence_balance_bonus(pattern)
    scored -= academic_support_penalty(pattern.members, final=True)
    return max(0.0, scored)


def feedback_adjustment_for_pattern(base_score: float, pattern: MeetingPattern) -> tuple[float, dict[str, Any]]:
    if FEEDBACK_LEARNER is None:
        return 0.0, {"feedback_count": 0, "word_memory_hits": 0}
    return FEEDBACK_LEARNER.adjustment(
        actors=pattern.members,
        topic=pattern.convening_theme,
        words=pattern_feedback_words(pattern),
        base_score=base_score,
        recommendation_id=pattern_recommendation_id(pattern),
        opportunity_id=opportunity_id_for_pattern(pattern),
    )


def neural_candidate_payload(base_score: float, pattern: MeetingPattern) -> dict[str, Any]:
    role_counts = actor_role_counts(pattern.members)
    actor_roles = [actor_role(member) for member in pattern.members]
    actor_subroles = [member_subrole_and_quota_weight(member)[0] for member in pattern.members]
    actor_types = {member: actor_role(member) for member in pattern.members}
    words = pattern_feedback_words(pattern)
    return {
        "recommendation_id": pattern_recommendation_id(pattern),
        "opportunity_id": opportunity_id_for_pattern(pattern),
        "opportunity_key": pattern.opportunity_key,
        "pair": " x ".join(pattern.members),
        "actors": list(pattern.members),
        "actor_types": actor_types,
        "actor_roles": actor_roles,
        "actor_subroles": [subrole for subrole in actor_subroles if subrole],
        "topic": pattern.convening_theme,
        "cluster": ", ".join(pattern.concrete_clusters),
        "cluster_label": ", ".join(pattern.concrete_clusters),
        "words": words,
        "reason": " ".join([
            pattern.reason,
            pattern.why_these_members,
            pattern.editorial_justification,
            pattern.possible_agenda,
        ]),
        "base_score": base_score,
        "score": pattern.score,
        "raw_score": pattern.score,
        "rule_final_score": base_score,
        "static_support_score": pattern.static_support_score,
        "live_signal_boost": pattern.live_signal_boost,
        "evidence_count": len(pattern.recent_signals),
        "live_event_count": pattern.live_event_count,
        "academic_evidence_count": role_counts["academic_support"],
        "company_evidence_count": role_counts["profit_anchor"] + role_counts["implementation_anchor"],
        "freshness_days": 0 if pattern.live_event_count else 365,
        "academic_flag": role_counts["academic_support"] > 0,
        "academic_only_flag": role_counts["academic_support"] == len(pattern.members),
        "company_need_confirmed": pattern_has_application_capability_bridge(pattern),
        "profit_anchor_count": role_counts["profit_anchor"],
        "implementation_anchor_count": role_counts["implementation_anchor"],
        "academic_support_count": role_counts["academic_support"],
        "context_actor_count": role_counts["context_actor"],
    }


def neural_feedback_adjustment_for_pattern(base_score: float, pattern: MeetingPattern) -> tuple[float, dict[str, Any]]:
    if NEURAL_FEEDBACK_RERANKER is None:
        return 0.0, {
            "nn_probability_useful": 0.5,
            "nn_delta": 0.0,
            "nn_model_examples": 0,
            "nn_max_delta": 0.0,
            "nn_model_status": "unavailable",
        }
    if (
        pattern.candidate_source == "static_excel"
        and low_score_for_pattern(pattern)
        and weak_evidence_for_pattern(pattern)
        and not pattern.legacy_candidate_id
        and not pattern.legacy_opportunity_key
        and base_score < MIN_FINAL_PAIR_SCORE
    ):
        return 0.0, {
            "nn_probability_useful": 0.5,
            "nn_delta": 0.0,
            "nn_model_examples": getattr(NEURAL_FEEDBACK_RERANKER, "n_examples", 0),
            "nn_max_delta": getattr(NEURAL_FEEDBACK_RERANKER, "max_delta", 0.0),
            "nn_model_status": "skipped_low_weak_full_universe_row",
        }
    result = NEURAL_FEEDBACK_RERANKER.score_delta(neural_candidate_payload(base_score, pattern))
    return float(result.get("nn_delta", 0.0) or 0.0), result


def macro_context_penalty(pattern: MeetingPattern) -> float:
    text = normalize_search_text(" ".join(pattern.recent_signals))
    macro_hits = normalized_hit_count(MACRO_TERMS | MARKET_NOISE_TERMS, text)
    problem_meta = weighted_problem_hit_score(text)
    action_meta = weighted_action_hit_score(text)
    problem_hits = problem_meta["score"]
    action_hits = action_meta["score"]
    rlp_hits = normalized_hit_count(["rheinland-pfalz", "rlp", "mainz", "trier", "koblenz", "pfalz", "bingen"], text)

    if macro_hits >= 2 and action_hits < 2:
        return 12.0
    if macro_hits >= 1 and problem_hits == 0 and rlp_hits == 0:
        return 8.0
    return 0.0


def has_operational_evidence(pattern: MeetingPattern) -> bool:
    text = normalize_search_text(" ".join(pattern.recent_signals))
    return (
        weighted_problem_hit_score(text)["score"] >= 1
        and weighted_action_hit_score(text)["score"] >= 1
        and weighted_problem_hit_score(text)["static"] >= 1
        and normalized_hit_count(MARKET_NOISE_TERMS, text) <= 1
    )


def critical_action_penalty(pattern: MeetingPattern) -> float:
    penalty = 0.0
    counts = actor_role_counts(pattern.members)
    evidence_count = len(pattern.recent_signals)
    distinct_members = distinct_signal_member_count(pattern.recent_signals)
    if evidence_count < 2:
        penalty += 6.0
    if distinct_members < 2:
        penalty += 5.0
    if counts["context_actor"] and not counts["academic_support"]:
        penalty += 7.0
    if counts["context_actor"] and not (counts["profit_anchor"] + counts["implementation_anchor"]):
        penalty += 8.0
    if not has_non_generic_zirp_question(pattern.shared_topics, pattern.bridge_topics):
        penalty += 5.0
    if not pattern.possible_agenda or len(pattern.possible_agenda) < 32:
        penalty += 4.0
    if not pattern_has_application_capability_bridge(pattern):
        penalty += 12.0
    if not has_valid_complementarity_sentence(pattern):
        penalty += 10.0
    penalty += abstract_overlap_penalty(
        pattern.shared_topics,
        pattern.bridge_topics,
        " ".join([
            pattern.convening_theme,
            pattern.why_these_members,
            pattern.editorial_justification,
            pattern.possible_agenda,
            " ".join(pattern.recent_signals),
        ]),
    )
    penalty += macro_context_penalty(pattern)
    penalty += pair_evidence_balance_penalty(pattern)
    return min(32.0, penalty)


def final_pattern_score(base_score: float, pattern: MeetingPattern) -> float:
    cached = getattr(pattern, "_cached_final_pattern_score", None)
    if cached is not None:
        return float(cached)
    if (
        pattern.candidate_source == "static_excel"
        and low_score_for_pattern(pattern)
        and weak_evidence_for_pattern(pattern)
        and not pattern.legacy_candidate_id
        and not pattern.legacy_opportunity_key
        and float(base_score) < MIN_FINAL_PAIR_SCORE
    ):
        value = min(100.0, max(0.0, float(base_score)))
        setattr(pattern, "_cached_final_pattern_score", value)
        return value
    rule_score = float(base_score) if pattern.candidate_source == "static_excel" else rule_pattern_score(base_score, pattern)
    learner_adjustment, _ = feedback_adjustment_for_pattern(rule_score, pattern)
    human_adjustment, _ = explicit_human_feedback_adjustment_for_pattern(pattern)
    neural_adjustment, _ = neural_feedback_adjustment_for_pattern(rule_score, pattern)
    value = min(100.0, max(0.0, rule_score + learner_adjustment + human_adjustment + neural_adjustment - critical_action_penalty(pattern)))
    setattr(pattern, "_cached_final_pattern_score", value)
    return value


def structural_bonus_amount(base_score: float, pattern: MeetingPattern) -> float:
    if pattern.candidate_source == "static_excel":
        return 0.0
    return round(rule_pattern_score(base_score, pattern) - base_score, 3)


def display_points_for_score(score: float) -> int:
    try:
        return max(0, min(100, int(round(float(score)))))
    except Exception:
        return 0


def public_status_for_points(points: int) -> str:
    if points >= 20:
        return "Hauptlinie"
    if points >= 15:
        return "Sondierungshypothese"
    if points >= 10:
        return "Watchlist"
    return "nicht öffentlich"


HIGH_GRAVITY_MEMBERS = {
    member
    for member in MEMBER_ENVIRONMENT_WEIGHTS
    if member_environment_score(member) >= 1.11
}

def lp_var(pattern: MeetingPattern) -> str:
    if pattern.scip_variable:
        return re.sub(r"[^A-Za-z0-9_]+", "_", pattern.scip_variable)
    return f"y{pattern.pattern_id}"


def member_role_and_quota_weight(member_name: str) -> tuple[str, float]:
    configured = STATIC_MEMBER_ROLE_WEIGHTS.get(normalize_search_text(member_name))
    if configured:
        return configured
    role = actor_role(member_name)
    return role, STATIC_ROLE_INVERSE_WEIGHTS.get(role, 0.0)


def role_mass_for_pattern(pattern: MeetingPattern, role: str) -> float:
    return sum(
        weight
        for member_role, weight in (member_role_and_quota_weight(member) for member in pattern.members)
        if member_role == role
    )


def member_subrole_and_quota_weight(member_name: str) -> tuple[str, float]:
    configured = STATIC_MEMBER_SUBROLE_WEIGHTS.get(normalize_search_text(member_name))
    if configured:
        return configured
    subrole = ""
    weight = 0.0
    key = normalize_search_text(member_name)
    if "hochschule" in key or "universitat" in key or "university" in key or "rptu" in key or "whu" in key:
        subrole = "university_applied_sciences" if "hochschule" in key and "universitat" not in key else "university_research"
        weight = STATIC_SUBROLE_INVERSE_WEIGHTS.get(subrole, 0.0)
    return subrole, weight


def subrole_mass_for_pattern(pattern: MeetingPattern, subrole: str) -> float:
    return sum(
        weight
        for member_subrole, weight in (member_subrole_and_quota_weight(member) for member in pattern.members)
        if member_subrole == subrole
    )


def academic_actor_slots_for_pattern(pattern: MeetingPattern) -> int:
    if pattern.academic_pair_flag and pattern.academic_academic_flag:
        return max(2, sum(1 for member in pattern.members if member_subrole_and_quota_weight(member)[0].startswith("university_")))
    if pattern.academic_pair_flag and not any(member_subrole_and_quota_weight(member)[0].startswith("university_") for member in pattern.members):
        return 1
    return sum(
        1
        for member in pattern.members
        if member_subrole_and_quota_weight(member)[0].startswith("university_")
    )


def has_academic_actor(pattern: MeetingPattern) -> bool:
    return bool(pattern.academic_pair_flag) or academic_actor_slots_for_pattern(pattern) > 0


def non_academic_pair_for_pattern(pattern: MeetingPattern) -> int:
    if pattern.non_academic_pair_flag:
        return 1
    return int(academic_actor_slots_for_pattern(pattern) == 0 and not has_academic_actor(pattern))


def weak_evidence_for_pattern(pattern: MeetingPattern) -> int:
    return int(pattern.weak_evidence_flag > 0)


def low_score_for_pattern(pattern: MeetingPattern) -> int:
    return int(pattern.low_score_flag > 0)


def academic_academic_for_pattern(pattern: MeetingPattern) -> int:
    if pattern.academic_academic_flag:
        return 1
    return int(academic_actor_slots_for_pattern(pattern) >= 2)


def geo_proximity_norm_for_pattern(pattern: MeetingPattern) -> float:
    if pattern.geo_proximity_norm > 0:
        return min(1.0, max(0.0, pattern.geo_proximity_norm))
    if pattern.geo_proximity_score > 1.0:
        return min(1.0, max(0.0, pattern.geo_proximity_score / 100.0))
    return min(1.0, max(0.0, pattern.geo_proximity_score))


def export_scip_incidence_matrix(patterns: list[MeetingPattern], path: Path) -> None:
    members = sorted({member for pattern in patterns for member in pattern.members})
    clusters = sorted({cluster for pattern in patterns for cluster in pattern.concrete_clusters})
    opportunities = sorted({opportunity_id_for_pattern(pattern) for pattern in patterns})
    rows = []
    for pattern in patterns:
        row: dict[str, Any] = {
            "candidate_id": pattern.candidate_id or str(pattern.pattern_id),
            "scip_variable": lp_var(pattern),
            "candidate_source": pattern.candidate_source,
            "members": " | ".join(pattern.members),
            "concrete_clusters": ", ".join(pattern.concrete_clusters),
            "opportunity_id": opportunity_id_for_pattern(pattern),
            "opportunity_key": pattern.opportunity_key,
            "member_pair_key": pattern.member_pair_key,
            "static_support_score": round(pattern.static_support_score, 3),
            "live_signal_boost": round(pattern.live_signal_boost, 3),
            "feedback_adjustment": round(
                feedback_adjustment_for_pattern(pattern.score, pattern)[0]
                + explicit_human_feedback_adjustment_for_pattern(pattern)[0],
                3,
            ),
            "neural_delta": round(neural_feedback_adjustment_for_pattern(pattern.score, pattern)[0], 3),
            "critical_action_penalty": round(critical_action_penalty(pattern), 3),
            "objective_score": round(final_pattern_score(pattern.score, pattern), 3),
        }
        for member in members:
            row["member__" + re.sub(r"[^A-Za-z0-9]+", "_", member)[:60]] = int(member in pattern.members)
        for cluster in clusters:
            row["cluster__" + re.sub(r"[^A-Za-z0-9]+", "_", cluster)[:60]] = int(cluster in pattern.concrete_clusters)
        opportunity = opportunity_id_for_pattern(pattern)
        for opportunity_id in opportunities:
            row["opportunity__" + re.sub(r"[^A-Za-z0-9]+", "_", opportunity_id)[:60]] = int(opportunity_id == opportunity)
        rows.append(row)
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")


def _write_lp_linear_constraint(
    f: Any,
    name: str,
    terms: list[str],
    sense: str,
    rhs: str | int | float,
    *,
    max_line_chars: int = 900,
) -> None:
    if not terms:
        return
    line = f" {name}: "
    first = True
    for term in terms:
        piece = term if first else f" + {term}"
        if len(line) + len(piece) > max_line_chars:
            f.write(line.rstrip() + "\n")
            line = "  " + (term if first else f"+ {term}")
        else:
            line += piece
        first = False
    line += f" {sense} {rhs}"
    f.write(line + "\n")


def write_scip_model(patterns: list[MeetingPattern], active_members: list[str], lp_path: Path) -> None:
    lp_path.parent.mkdir(parents=True, exist_ok=True)

    with lp_path.open("w", encoding="utf-8", newline="\n") as f:
        f.write("\\ ZIRP meeting pattern optimization\n")
        f.write("Maximize\n")
        terms = [
            f"{final_pattern_score(pattern.score, pattern):.4f} {lp_var(pattern)}"
            for pattern in patterns
        ]
        f.write(" obj: " + (" + ".join(terms) if terms else "0") + "\n")
        f.write("Subject To\n")

        if patterns:
            _write_lp_linear_constraint(
                f,
                "selected_meetings",
                [lp_var(pattern) for pattern in patterns],
                "=",
                MAX_MEETINGS,
            )

        for member in active_members:
            member_patterns = [lp_var(pattern) for pattern in patterns if member in pattern.members]
            if member_patterns:
                member_limit = STATIC_MEMBER_LIMITS.get(normalize_search_text(member), MAX_MEETINGS_PER_MEMBER)
                cname = "member_" + re.sub(r"[^A-Za-z0-9]+", "_", member)[:50]
                _write_lp_linear_constraint(f, cname, member_patterns, "<=", member_limit)

        # Limit how often the same concrete cluster appears in the selected portfolio.
        cluster_to_patterns: dict[str, list[str]] = defaultdict(list)
        for pattern in patterns:
            for cluster in pattern.concrete_clusters:
                cluster_to_patterns[cluster].append(lp_var(pattern))

        for cluster, vars_ in cluster_to_patterns.items():
            cluster_limit = STATIC_CLUSTER_LIMITS.get(cluster, MAX_SELECTED_PER_CLUSTER)
            if len(vars_) <= cluster_limit:
                continue
            cname = "cluster_" + re.sub(r"[^A-Za-z0-9]+", "_", cluster)[:50]
            _write_lp_linear_constraint(f, cname, vars_, "<=", cluster_limit)

        # Clique-style tightening: alternate rows for the same opportunity line
        # should not crowd out distinct matches.
        opportunity_to_patterns: dict[str, list[str]] = defaultdict(list)
        for pattern in patterns:
            opportunity_to_patterns[opportunity_id_for_pattern(pattern)].append(lp_var(pattern))

        for opportunity_id, vars_ in opportunity_to_patterns.items():
            if len(vars_) <= MAX_SELECTED_PER_OPPORTUNITY:
                continue
            cname = "opportunity_" + re.sub(r"[^A-Za-z0-9]+", "_", opportunity_id)[:50]
            _write_lp_linear_constraint(f, cname, vars_, "<=", MAX_SELECTED_PER_OPPORTUNITY)

        for role, limit in sorted(STATIC_ROLE_MASS_LIMITS.items()):
            terms = []
            for pattern in patterns:
                coeff = role_mass_for_pattern(pattern, role)
                if coeff > 0:
                    terms.append(f"{coeff:.6f} {lp_var(pattern)}")
            if terms:
                cname = "role_quota_" + re.sub(r"[^A-Za-z0-9]+", "_", role)[:50]
                _write_lp_linear_constraint(f, cname, terms, "<=", f"{limit:.6f}")

        for subrole, limit in sorted(STATIC_SUBROLE_MASS_LIMITS.items()):
            terms = []
            for pattern in patterns:
                coeff = subrole_mass_for_pattern(pattern, subrole)
                if coeff > 0:
                    terms.append(f"{coeff:.6f} {lp_var(pattern)}")
            if terms:
                cname = "subrole_quota_" + re.sub(r"[^A-Za-z0-9]+", "_", subrole)[:50]
                _write_lp_linear_constraint(f, cname, terms, "<=", f"{limit:.6f}")

        academic_match_vars = [lp_var(pattern) for pattern in patterns if has_academic_actor(pattern)]
        if academic_match_vars and STATIC_MAX_ACADEMIC_MATCHES >= 0:
            _write_lp_linear_constraint(f, "academic_match_cap", academic_match_vars, "<=", STATIC_MAX_ACADEMIC_MATCHES)

        academic_slot_terms = [
            f"{academic_actor_slots_for_pattern(pattern)} {lp_var(pattern)}"
            for pattern in patterns
            if academic_actor_slots_for_pattern(pattern) > 0
        ]
        if academic_slot_terms and STATIC_MAX_ACADEMIC_ACTOR_SLOTS >= 0:
            _write_lp_linear_constraint(
                f,
                "academic_actor_slot_cap",
                academic_slot_terms,
                "<=",
                STATIC_MAX_ACADEMIC_ACTOR_SLOTS,
            )

        non_academic_terms = [
            lp_var(pattern)
            for pattern in patterns
            if non_academic_pair_for_pattern(pattern) > 0
        ]
        if non_academic_terms and STATIC_MIN_NON_ACADEMIC_PAIRS > 0:
            _write_lp_linear_constraint(
                f,
                "non_academic_pair_floor",
                non_academic_terms,
                ">=",
                STATIC_MIN_NON_ACADEMIC_PAIRS,
            )

        weak_evidence_terms = [
            lp_var(pattern)
            for pattern in patterns
            if weak_evidence_for_pattern(pattern) > 0
        ]
        if weak_evidence_terms and STATIC_MAX_WEAK_EVIDENCE_MATCHES >= 0:
            _write_lp_linear_constraint(
                f,
                "weak_evidence_cap",
                weak_evidence_terms,
                "<=",
                STATIC_MAX_WEAK_EVIDENCE_MATCHES,
            )

        low_score_terms = [
            lp_var(pattern)
            for pattern in patterns
            if low_score_for_pattern(pattern) > 0
        ]
        if low_score_terms and STATIC_MAX_LOW_SCORE_MATCHES >= 0:
            _write_lp_linear_constraint(f, "low_score_cap", low_score_terms, "<=", STATIC_MAX_LOW_SCORE_MATCHES)

        geo_terms = [
            f"{geo_proximity_norm_for_pattern(pattern):.6f} {lp_var(pattern)}"
            for pattern in patterns
            if geo_proximity_norm_for_pattern(pattern) > 0
        ]
        if geo_terms and STATIC_MIN_GEO_PROXIMITY_SUM > 0:
            _write_lp_linear_constraint(
                f,
                "geo_proximity_floor",
                geo_terms,
                ">=",
                f"{STATIC_MIN_GEO_PROXIMITY_SUM:.6f}",
            )

        academic_academic_terms = [
            lp_var(pattern)
            for pattern in patterns
            if academic_academic_for_pattern(pattern) > 0
        ]
        if academic_academic_terms and STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS >= 0:
            _write_lp_linear_constraint(
                f,
                "academic_academic_cap",
                academic_academic_terms,
                "<=",
                STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS,
            )

        f.write("Binary\n")
        for pattern in patterns:
            f.write(f" {lp_var(pattern)}\n")
        f.write("End\n")


def run_scip(lp_path: Path) -> tuple[set[int], str]:
    scip_exe = find_scip_exe()
    cmd = f'read "{lp_path}" optimize display solution quit'
    proc = subprocess.run(
        [str(scip_exe), "-c", cmd],
        cwd=str(SCIP_ROOT),
        capture_output=True,
        text=True,
    )
    output = (proc.stdout or "") + "\n" + (proc.stderr or "")
    selected = set()
    for var, value in re.findall(r"\b([A-Za-z][A-Za-z0-9_]*)\s+([01](?:\.0+)?)\b", output):
        if float(value) <= 0.5:
            continue
        if var.startswith("y") and var[1:].isdigit():
            selected.add(int(var[1:]))
            continue
        match = re.search(r"(\d+)$", var)
        if match:
            selected.add(int(match.group(1)))
    return selected, output


def greedy_fallback(patterns: list[MeetingPattern]) -> set[int]:
    usage = Counter()
    used_clusters = Counter()
    used_opportunities = Counter()
    used_role_mass = Counter()
    used_subrole_mass = Counter()
    used_academic_matches = 0
    used_academic_actor_slots = 0
    used_non_academic_pairs = 0
    used_weak_evidence_matches = 0
    used_low_score_matches = 0
    used_geo_proximity_sum = 0.0
    used_academic_academic_pairs = 0
    selected: set[int] = set()

    for pattern in sorted(patterns, key=lambda p: (-final_pattern_score(p.score, p), len(p.members))):
        if len(selected) >= MAX_MEETINGS:
            break

        if len(pattern.members) != 2:
            continue

        if final_pattern_score(pattern.score, pattern) < MIN_FINAL_PAIR_SCORE:
            continue

        if any(usage[member] >= STATIC_MEMBER_LIMITS.get(normalize_search_text(member), MAX_MEETINGS_PER_MEMBER) for member in pattern.members):
            continue

        if any(used_clusters[cluster] >= STATIC_CLUSTER_LIMITS.get(cluster, MAX_SELECTED_PER_CLUSTER) for cluster in pattern.concrete_clusters):
            continue

        opportunity_id = opportunity_id_for_pattern(pattern)
        if used_opportunities[opportunity_id] >= MAX_SELECTED_PER_OPPORTUNITY:
            continue

        role_masses = {
            role: role_mass_for_pattern(pattern, role)
            for role in STATIC_ROLE_MASS_LIMITS
        }
        if any(used_role_mass[role] + mass > STATIC_ROLE_MASS_LIMITS[role] + 1e-9 for role, mass in role_masses.items()):
            continue

        subrole_masses = {
            subrole: subrole_mass_for_pattern(pattern, subrole)
            for subrole in STATIC_SUBROLE_MASS_LIMITS
        }
        if any(used_subrole_mass[subrole] + mass > STATIC_SUBROLE_MASS_LIMITS[subrole] + 1e-9 for subrole, mass in subrole_masses.items()):
            continue

        academic_slots = academic_actor_slots_for_pattern(pattern)
        if academic_slots:
            if used_academic_matches + 1 > STATIC_MAX_ACADEMIC_MATCHES:
                continue
            if used_academic_actor_slots + academic_slots > STATIC_MAX_ACADEMIC_ACTOR_SLOTS:
                continue

        non_academic_pair = non_academic_pair_for_pattern(pattern)
        weak_evidence = weak_evidence_for_pattern(pattern)
        low_score = low_score_for_pattern(pattern)
        academic_academic = academic_academic_for_pattern(pattern)
        geo_score = geo_proximity_norm_for_pattern(pattern)

        if used_weak_evidence_matches + weak_evidence > STATIC_MAX_WEAK_EVIDENCE_MATCHES:
            continue
        if used_low_score_matches + low_score > STATIC_MAX_LOW_SCORE_MATCHES:
            continue
        if used_academic_academic_pairs + academic_academic > STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS:
            continue

        remaining_after = MAX_MEETINGS - (len(selected) + 1)
        needed_non_academic_after = max(
            0,
            STATIC_MIN_NON_ACADEMIC_PAIRS - (used_non_academic_pairs + non_academic_pair),
        )
        if needed_non_academic_after > remaining_after:
            continue
        if used_geo_proximity_sum + geo_score + remaining_after < STATIC_MIN_GEO_PROXIMITY_SUM:
            continue

        selected.add(pattern.pattern_id)

        for member in pattern.members:
            usage[member] += 1

        for cluster in pattern.concrete_clusters:
            used_clusters[cluster] += 1

        used_opportunities[opportunity_id] += 1
        for role, mass in role_masses.items():
            used_role_mass[role] += mass
        for subrole, mass in subrole_masses.items():
            used_subrole_mass[subrole] += mass
        if academic_slots:
            used_academic_matches += 1
            used_academic_actor_slots += academic_slots
        used_non_academic_pairs += non_academic_pair
        used_weak_evidence_matches += weak_evidence
        used_low_score_matches += low_score
        used_geo_proximity_sum += geo_score
        used_academic_academic_pairs += academic_academic

    return selected


def export_profiles(profiles: dict[str, MemberProfile], path: Path) -> None:
    rows = []
    for profile in sorted(profiles.values(), key=lambda p: (p.member_id, p.name)):
        rows.append({
            "member_id": profile.member_id,
            "mitglied": profile.name,
            "homepage": profile.homepage,
            "event_count": profile.event_count,
            "top_topics": ", ".join(display_topic(t) for t, _ in profile.topics.most_common(8)),
            "top_clusters": ", ".join(display_topic(t) for t, _ in profile.clusters.most_common(6)),
            "top_sections": ", ".join(t for t, _ in profile.sections.most_common(5)),
            "relevant_names": ", ".join(n for n, _ in profile.names.most_common(8)),
            "avg_decision_score": round(profile.avg_decision_score, 2),
            "avg_event_score": round(profile.avg_event_score, 2),
        })
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")


def export_patterns(patterns: list[MeetingPattern], selected_ids: set[int], path: Path) -> pd.DataFrame:
    rows = []
    columns = [
        "selected", "pattern_id", "candidate_id", "scip_variable", "candidate_source",
        "recommendation_id", "opportunity_id", "opportunity_key", "member_pair_key", "score", "display_points",
        "public_status", "raw_score", "static_support_score", "live_signal_boost",
        "live_event_count", "live_evidence_status", "live_signal_titles", "live_signal_urls",
        "geo_proximity_score", "geo_proximity_norm", "distance_km", "same_geo_region",
        "geo_score_reason", "candidate_status", "filter_reason",
        "weak_evidence_flag", "low_score_flag", "academic_pair_flag",
        "academic_academic_flag", "non_academic_pair_flag",
        "legacy_candidate_id", "legacy_opportunity_key", "old_candidate_max_score", "old_candidate_rows",
        "rule_final_score",
        "final_score", "structural_bonus", "memory_adjustment", "nn_adjustment",
        "neural_probability_useful", "neural_delta", "neural_model_examples",
        "neural_max_delta", "neural_model_status",
        "human_feedback_adjustment", "feedback_adjustment", "total_learning_adjustment",
        "feedback_count", "human_feedback_count",
        "human_feedback_exact_count", "human_feedback_similar_count",
        "human_feedback_label_adjustment", "human_feedback_reason_adjustment",
        "human_feedback_target_adjustment", "human_feedback_selection_adjustment",
        "human_feedback_positive_count", "human_feedback_negative_count",
        "human_feedback_top_labels", "human_feedback_top_reasons",
        "feedback_evidence_count", "word_memory_hits", "critical_action_penalty",
        "pair_evidence_balance", "pair_evidence_balance_exception",
        "role_complementarity_score", "application_capability_bridge", "valid_complementarity_sentence",
        "environment_multiplier", "cluster_relevance_multiplier", "academic_support_count",
        "profit_anchor_count", "implementation_anchor_count", "context_actor_count",
        "cluster_size", "members", "actor_use_modes",
        "shared_topics", "bridge_topics", "concrete_clusters", "cluster_problem", "relevant_names", "reason",
        "convening_theme", "convening_typology", "convening_stage", "purpose_maturity",
        "typology_confidence", "typology_reason", "stage_gate_status", "stage_gate_reason",
        "north_star_purpose", "participants_logic", "required_inputs", "expected_output", "expected_outputs",
        "stage_allowed_outputs", "stage_forbidden_claims", "participant_takeaway", "success_metrics",
        "next_decision_gate", "next_engagement_move",
        "required_roles", "missing_typology_roles", "do_not_ask_for",
        "why_these_members", "editorial_justification",
        "empfehlung", "zirp_fit", "aktionsreife", "signalstaerke",
        "card_title", "decision_summary", "decision_line", "risk_line", "next_line",
        "why_now", "shared_tension", "ideal_guests", "recommended_format", "possible_output",
        "expected_output_model", "commitment_test", "complementarity_reason_1", "complementarity_reason_2",
        "disqualifying_uncertainty", "non_selection_reason",
        "next_best_action", "recommendation_limit", "maturity_level",
        "why_this_actor", "why_this_counterpart", "concrete_joint_action",
        "evidence_quality", "sufficiency", "consequence", "uncertainty",
        "action_decision", "critical_decision_reason",
        "rlp_relevance", "recent_signals", "suggested_format", "possible_agenda",
    ]
    for pattern in patterns:
        raw_score = float(pattern.score)
        rule_final_score = raw_score if pattern.candidate_source == "static_excel" else rule_pattern_score(raw_score, pattern)
        memory_adjustment, feedback_meta = feedback_adjustment_for_pattern(rule_final_score, pattern)
        human_feedback_adjustment, human_feedback_meta = explicit_human_feedback_adjustment_for_pattern(pattern)
        neural_adjustment, neural_meta = neural_feedback_adjustment_for_pattern(rule_final_score, pattern)
        feedback_adjustment_total = memory_adjustment + human_feedback_adjustment
        total_learning_adjustment = feedback_adjustment_total + neural_adjustment
        critical_penalty = critical_action_penalty(pattern)
        final_score_value = final_pattern_score(pattern.score, pattern)
        environment_multiplier = pattern_environment_multiplier(pattern)
        cluster_relevance_multiplier = pattern_cluster_relevance_multiplier(pattern)
        structural_bonus = rule_final_score - raw_score
        role_counts = actor_role_counts(pattern.members)
        evidence_count = len(pattern.recent_signals)
        distinct_evidence_members = distinct_signal_member_count(pattern.recent_signals)
        common_problem_quality = "stark" if pattern.shared_topics and pattern.bridge_topics else ("mittel" if pattern.shared_topics or pattern.bridge_topics else "schwach")
        rlp_fit = rlp_fit_for_pattern(pattern.members, pattern.shared_topics, pattern.bridge_topics, pattern.recent_signals)
        non_generic_question = has_non_generic_zirp_question(pattern.shared_topics, pattern.bridge_topics)
        balance = pair_evidence_balance(pattern)
        balance_exception = strong_problem_owner_capability_exception(pattern)
        labels = classify_convening(
            final_score_value,
            evidence_count,
            rlp_fit,
            common_problem_quality,
            distinct_evidence_members,
            non_generic_question,
            pair_balance=balance,
            pair_balance_exception=balance_exception,
        )
        judgment = critical_judgment_for_pattern(
            pattern,
            labels,
            evidence_count,
            distinct_evidence_members,
            rlp_fit,
            common_problem_quality,
        )
        suggested_format = convening_format_for_label(labels["empfehlung"], pattern.suggested_format)
        missing_actors = missing_actors_for_topics(pattern.shared_topics, pattern.bridge_topics)
        next_action = next_action_for_label(labels["empfehlung"], pattern.members, pattern.possible_agenda)
        recommendation_limit = recommendation_limit_for_label(labels["empfehlung"])
        has_coalition = has_existing_coalition_signal(pattern)
        has_resource = has_resource_signal(pattern)
        has_launch = has_launch_or_publication_signal(pattern)
        raw_typology = infer_convening_typology(
            evidence_count,
            distinct_evidence_members,
            common_problem_quality,
            labels["aktionsreife"],
            has_existing_coalition=has_coalition,
            has_resource_signal_value=has_resource,
            has_launch_or_publication_signal_value=has_launch,
        )
        clear_problem = has_clear_convening_problem(pattern)
        member_role_set = set().union(*(infer_member_roles(member) for member in pattern.members))
        typology, stage_gate_reason = apply_typology_stage_gate(
            raw_typology,
            evidence_count=evidence_count,
            distinct_evidence_members=distinct_evidence_members,
            common_problem_quality=common_problem_quality,
            action_readiness=labels["aktionsreife"],
            has_clear_problem=clear_problem,
            member_roles=member_role_set,
            has_resource_signal_value=has_resource,
            has_launch_or_publication_signal_value=has_launch,
        )
        stage_gate_status = "pass" if not stage_gate_reason else "downgraded"
        purpose_maturity = infer_purpose_maturity(
            evidence_count,
            distinct_evidence_members,
            common_problem_quality,
            labels["aktionsreife"],
            has_coalition,
            has_resource,
        )
        stage = stage_for_typology(typology)
        missing_typology_roles = missing_roles_for_typology(member_role_set, typology)
        next_move = next_engagement_move_for_stage(
            stage,
            missing_roles=missing_typology_roles,
            has_clear_problem=clear_problem,
            member_count=len(pattern.members),
        )
        design_arc = design_arc_for_pattern(pattern, typology, stage, next_move, TYPOLOGY_REQUIRED_ROLES.get(typology, set()), missing_typology_roles)
        next_action = next_move.value
        if missing_typology_roles:
            next_action = recommended_action_for_typology(typology, judgment["evidence_quality"], missing_typology_roles)
        display_points = display_points_for_score(final_score_value)
        rows.append({
            "selected": int(pattern.pattern_id in selected_ids),
            "pattern_id": pattern.pattern_id,
            "candidate_id": pattern.candidate_id,
            "scip_variable": lp_var(pattern),
            "candidate_source": pattern.candidate_source,
            "recommendation_id": pattern_recommendation_id(pattern),
            "opportunity_id": opportunity_id_for_pattern(pattern),
            "opportunity_key": pattern.opportunity_key,
            "member_pair_key": pattern.member_pair_key,
            "score": round(final_score_value, 3),
            "display_points": display_points,
            "public_status": public_status_for_points(display_points),
            "raw_score": round(raw_score, 3),
            "static_support_score": round(pattern.static_support_score, 3),
            "live_signal_boost": round(pattern.live_signal_boost, 3),
            "live_event_count": pattern.live_event_count,
            "live_evidence_status": pattern.live_evidence_status,
            "live_signal_titles": " | ".join(pattern.live_signal_titles),
            "live_signal_urls": " | ".join(pattern.live_signal_urls),
            "geo_proximity_score": round(pattern.geo_proximity_score, 3),
            "geo_proximity_norm": round(geo_proximity_norm_for_pattern(pattern), 4),
            "distance_km": round(pattern.distance_km, 2),
            "same_geo_region": int(pattern.same_geo_region),
            "geo_score_reason": pattern.geo_score_reason,
            "candidate_status": pattern.candidate_status,
            "filter_reason": pattern.filter_reason,
            "weak_evidence_flag": weak_evidence_for_pattern(pattern),
            "low_score_flag": low_score_for_pattern(pattern),
            "academic_pair_flag": int(has_academic_actor(pattern)),
            "academic_academic_flag": academic_academic_for_pattern(pattern),
            "non_academic_pair_flag": non_academic_pair_for_pattern(pattern),
            "legacy_candidate_id": pattern.legacy_candidate_id,
            "legacy_opportunity_key": pattern.legacy_opportunity_key,
            "old_candidate_max_score": round(pattern.old_candidate_max_score, 3),
            "old_candidate_rows": pattern.old_candidate_rows,
            "rule_final_score": round(rule_final_score, 3),
            "final_score": round(final_score_value, 3),
            "structural_bonus": round(structural_bonus, 3),
            "memory_adjustment": round(memory_adjustment, 3),
            "nn_adjustment": round(neural_adjustment, 3),
            "neural_probability_useful": round(float(neural_meta.get("nn_probability_useful", 0.5) or 0.5), 4),
            "neural_delta": round(neural_adjustment, 3),
            "neural_model_examples": int(neural_meta.get("nn_model_examples", 0) or 0),
            "neural_max_delta": round(float(neural_meta.get("nn_max_delta", 0.0) or 0.0), 3),
            "neural_model_status": neural_meta.get("nn_model_status", ""),
            "human_feedback_adjustment": round(human_feedback_adjustment, 3),
            "feedback_adjustment": round(feedback_adjustment_total, 3),
            "total_learning_adjustment": round(total_learning_adjustment, 3),
            "feedback_count": int(feedback_meta.get("feedback_count", 0) or 0),
            "human_feedback_count": int(human_feedback_meta.get("human_feedback_count", 0) or 0),
            "human_feedback_exact_count": int(human_feedback_meta.get("human_feedback_exact_count", 0) or 0),
            "human_feedback_similar_count": int(human_feedback_meta.get("human_feedback_similar_count", 0) or 0),
            "human_feedback_label_adjustment": round(float(human_feedback_meta.get("human_feedback_label_adjustment", 0.0) or 0.0), 3),
            "human_feedback_reason_adjustment": round(float(human_feedback_meta.get("human_feedback_reason_adjustment", 0.0) or 0.0), 3),
            "human_feedback_target_adjustment": round(float(human_feedback_meta.get("human_feedback_target_adjustment", 0.0) or 0.0), 3),
            "human_feedback_selection_adjustment": round(float(human_feedback_meta.get("human_feedback_selection_adjustment", 0.0) or 0.0), 3),
            "human_feedback_positive_count": int(human_feedback_meta.get("human_feedback_positive_count", 0) or 0),
            "human_feedback_negative_count": int(human_feedback_meta.get("human_feedback_negative_count", 0) or 0),
            "human_feedback_top_labels": human_feedback_meta.get("human_feedback_top_labels", ""),
            "human_feedback_top_reasons": human_feedback_meta.get("human_feedback_top_reasons", ""),
            "feedback_evidence_count": int(feedback_meta.get("feedback_evidence_count", 0) or 0),
            "word_memory_hits": int(feedback_meta.get("word_memory_hits", 0) or 0),
            "critical_action_penalty": round(critical_penalty, 3),
            "pair_evidence_balance": balance,
            "pair_evidence_balance_exception": int(balance_exception),
            "role_complementarity_score": round(pattern_role_complementarity_score(pattern), 3),
            "application_capability_bridge": int(pattern_has_application_capability_bridge(pattern)),
            "valid_complementarity_sentence": int(has_valid_complementarity_sentence(pattern)),
            "environment_multiplier": round(environment_multiplier, 4),
            "cluster_relevance_multiplier": round(cluster_relevance_multiplier, 4),
            "academic_support_count": role_counts["academic_support"],
            "profit_anchor_count": role_counts["profit_anchor"],
            "implementation_anchor_count": role_counts["implementation_anchor"],
            "context_actor_count": role_counts["context_actor"],
            "cluster_size": len(pattern.members),
            "members": " | ".join(ordered_members_for_display(pattern.members)),
            "actor_use_modes": " | ".join(f"{member}: {actor_use_mode_for_member(member)}" for member in ordered_members_for_display(pattern.members)),
            "shared_topics": ", ".join(display_topic(t) for t in pattern.shared_topics),
            "bridge_topics": ", ".join(display_topic(t) for t in pattern.bridge_topics),
            "concrete_clusters": ", ".join(display_topic(t) for t in pattern.concrete_clusters),
            "cluster_problem": " | ".join(str(cluster_info(t).get("problem", "")) for t in pattern.concrete_clusters if cluster_info(t)),
            "relevant_names": ", ".join(pattern.relevant_names),
            "reason": pattern.reason,
            "convening_theme": pattern.convening_theme,
            "convening_typology": typology.value,
            "convening_stage": stage.value,
            "purpose_maturity": purpose_maturity.value,
            "typology_confidence": typology_confidence(typology, evidence_count, distinct_evidence_members, clear_problem),
            "typology_reason": typology_reason_for_pattern(typology, evidence_count, distinct_evidence_members, common_problem_quality, labels["aktionsreife"]),
            "stage_gate_status": stage_gate_status,
            "stage_gate_reason": stage_gate_reason,
            "north_star_purpose": design_arc.north_star_purpose,
            "participants_logic": design_arc.participants_logic,
            "required_inputs": ", ".join(design_arc.required_inputs),
            "expected_output": typology_expected_output(typology),
            "expected_outputs": ", ".join(design_arc.expected_outputs),
            "stage_allowed_outputs": ", ".join(sorted(STAGE_OUTPUT_RULES.get(stage, {}).get("allowed_outputs", set()))),
            "stage_forbidden_claims": ", ".join(sorted(STAGE_OUTPUT_RULES.get(stage, {}).get("forbidden_claims", set()))),
            "participant_takeaway": design_arc.participant_takeaway,
            "success_metrics": ", ".join(design_arc.success_metrics),
            "next_decision_gate": design_arc.next_decision_gate,
            "next_engagement_move": next_move.value,
            "required_roles": ", ".join(sorted(TYPOLOGY_REQUIRED_ROLES.get(typology, set()))),
            "missing_typology_roles": ", ".join(sorted(missing_typology_roles)),
            "do_not_ask_for": typology_do_not_ask_for(typology),
            "why_these_members": pattern.why_these_members,
            "editorial_justification": pattern.editorial_justification,
            "empfehlung": labels["empfehlung"],
            "zirp_fit": labels["zirp_fit"],
            "aktionsreife": labels["aktionsreife"],
            "signalstaerke": labels["signalstaerke"],
            "card_title": "",
            "decision_summary": "",
            "decision_line": "",
            "risk_line": "",
            "next_line": "",
            "why_now": judgment["why_now"],
            "shared_tension": shared_tension_for_pattern(pattern.shared_topics, pattern.bridge_topics),
            "ideal_guests": missing_actors,
            "recommended_format": suggested_format,
            "possible_output": judgment["concrete_joint_action"],
            "expected_output_model": "",
            "commitment_test": "",
            "complementarity_reason_1": "",
            "complementarity_reason_2": "",
            "disqualifying_uncertainty": "",
            "non_selection_reason": "",
            "next_best_action": next_action,
            "recommendation_limit": recommendation_limit,
            "maturity_level": maturity_level_for_label(labels["empfehlung"]),
            "why_this_actor": judgment["why_this_actor"],
            "why_this_counterpart": judgment["why_this_counterpart"],
            "concrete_joint_action": judgment["concrete_joint_action"],
            "evidence_quality": judgment["evidence_quality"],
            "sufficiency": judgment["sufficiency"],
            "consequence": judgment["consequence"],
            "uncertainty": judgment["uncertainty"],
            "action_decision": judgment["action_decision"],
            "critical_decision_reason": judgment["critical_decision_reason"],
            "rlp_relevance": rlp_fit,
            "recent_signals": " || ".join(pattern.recent_signals),
            "suggested_format": suggested_format,
            "possible_agenda": pattern.possible_agenda,
        })
    df = pd.DataFrame(rows, columns=columns)
    if not df.empty:
        df = df.sort_values(["selected", "score"], ascending=[False, False])
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return df


def split_leaderboard_cell(value: Any) -> list[str]:
    text = clean_display_value(value)
    if not text:
        return []
    return [piece.strip() for piece in re.split(r"\s*[,|]\s*", text) if piece.strip()]


def explain_selection_outcomes(patterns_df: pd.DataFrame) -> pd.DataFrame:
    if patterns_df.empty:
        return patterns_df.copy()

    df = patterns_df.copy()
    df["score_num"] = pd.to_numeric(df.get("score", 0), errors="coerce").fillna(0)
    df = df.sort_values("score_num", ascending=False).reset_index(drop=True)
    df["rank_all_pairs"] = range(1, len(df) + 1)

    selected = df[df.get("selected", 0).astype(str) == "1"].copy()
    selected_members = Counter(
        member
        for value in selected.get("members", [])
        for member in split_leaderboard_cell(value)
    )
    selected_clusters = Counter(
        cluster
        for value in selected.get("concrete_clusters", [])
        for cluster in split_leaderboard_cell(value)
    )
    selected_opportunities = Counter(clean_display_value(value) for value in selected.get("opportunity_id", []))

    reasons = []
    selected_reasons = []
    blocked_members_col = []
    blocked_clusters_col = []
    blocked_opportunities_col = []

    for _, row in df.iterrows():
        is_selected = str(row.get("selected", "0")) == "1"
        members = split_leaderboard_cell(row.get("members", ""))
        clusters = split_leaderboard_cell(row.get("concrete_clusters", ""))
        opportunity_id = clean_display_value(row.get("opportunity_id", ""))

        blocked_members = [member for member in members if selected_members.get(member, 0)]
        blocked_clusters = [
            cluster for cluster in clusters
            if selected_clusters.get(cluster, 0) >= MAX_SELECTED_PER_CLUSTER
        ]
        blocked_opportunity = opportunity_id if selected_opportunities.get(opportunity_id, 0) and not is_selected else ""

        if is_selected:
            reason = ""
            selected_reason = (
                "Von SCIP ausgewählt, weil diese Linie im aktuellen Signalbild besonders tragfähig wirkt: "
                "klare Rollen, belastbarer Anlass und hoher erwarteter Nutzen für eine mögliche ZIRP-Sondierung."
            )
        else:
            selected_reason = ""
            parts = []
            if blocked_members:
                parts.append("blocked by selected member: " + ", ".join(blocked_members))
            if blocked_clusters:
                parts.append("cluster cap reached: " + ", ".join(blocked_clusters))
            if blocked_opportunity:
                parts.append("duplicate opportunity line")
            if not parts:
                parts.append("lower portfolio value under SCIP constraints")
            reason = " | ".join(parts)

        reasons.append(reason)
        selected_reasons.append(selected_reason)
        blocked_members_col.append(", ".join(blocked_members))
        blocked_clusters_col.append(", ".join(blocked_clusters))
        blocked_opportunities_col.append(blocked_opportunity)

    df["why_selected"] = selected_reasons
    df["why_not_selected"] = reasons
    df["non_selection_reason"] = reasons
    df["blocked_by_member"] = blocked_members_col
    df["blocked_by_cluster"] = blocked_clusters_col
    df["blocked_by_opportunity"] = blocked_opportunities_col
    return df.drop(columns=["score_num"])


def export_selection_leaderboard(patterns_df: pd.DataFrame, path: Path) -> pd.DataFrame:
    leaderboard = explain_selection_outcomes(patterns_df)
    leaderboard.to_csv(path, index=False, encoding="utf-8-sig")
    return leaderboard


def write_model_summary(
        path: Path,
        *,
        timestamp: str,
        patterns: list[MeetingPattern],
        selected_ids: set[int],
        solver_output: str,
) -> None:
    solver_used = "greedy_fallback" if "used greedy fallback" in solver_output.lower() else "SCIP"
    neural_summary = {
        "enabled": USE_NEURAL_FEEDBACK_RANKER,
        "available": NEURAL_FEEDBACK_RERANKER is not None,
        "status": getattr(NEURAL_FEEDBACK_RERANKER, "status", "unavailable") if NEURAL_FEEDBACK_RERANKER is not None else "unavailable",
        "model_examples": getattr(NEURAL_FEEDBACK_RERANKER, "n_examples", 0) if NEURAL_FEEDBACK_RERANKER is not None else 0,
        "max_delta": getattr(NEURAL_FEEDBACK_RERANKER, "max_delta", 0.0) if NEURAL_FEEDBACK_RERANKER is not None else 0.0,
        "dynamic_max_delta": (
            NEURAL_FEEDBACK_RERANKER.dynamic_max_delta()
            if NEURAL_FEEDBACK_RERANKER is not None
            else 0.0
        ),
    }
    summary = {
        "run_id": timestamp,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "model_name": "SCIP Match Selector",
        "model_type": "binary integer optimization",
        "mathematical_framing": "maximum-weight set-packing / strategic opportunity matching",
        "network": NETWORK_NAME,
        "network_config_dir": NETWORK_CONFIG_DIR,
        "static_optimizer_path": str(STATIC_OPTIMIZER_PATH),
        "scip_input_sheet": STATIC_SCIP_INPUT_SHEET,
        "frontier_sheet": STATIC_FRONTIER_SHEET,
        "full_universe_sheet": STATIC_PAIR_VECTOR_SHEET,
        "card_sheet": STATIC_CARD_SHEET,
        "optimizer_config": NETWORK_OPTIMIZER,
        "solver": solver_used,
        "scip_executable": str(find_scip_exe()) if solver_used == "SCIP" else "",
        "selection_rule": "maximize total final match score under compatibility, diversity and capacity constraints",
        "pair_only": PAIR_ONLY,
        "min_cluster_size": MIN_CLUSTER_SIZE,
        "max_cluster_size": MAX_CLUSTER_SIZE,
        "max_matches": MAX_MEETINGS,
        "max_public_cards": MAX_PUBLIC_CARDS,
        "max_matches_per_member": MAX_MEETINGS_PER_MEMBER,
        "max_selected_per_cluster": MAX_SELECTED_PER_CLUSTER,
        "max_selected_per_opportunity": MAX_SELECTED_PER_OPPORTUNITY,
        "role_quota_limits": STATIC_ROLE_MASS_LIMITS,
        "subrole_quota_limits": STATIC_SUBROLE_MASS_LIMITS,
        "max_academic_matches": STATIC_MAX_ACADEMIC_MATCHES,
        "max_academic_actor_slots": STATIC_MAX_ACADEMIC_ACTOR_SLOTS,
        "min_non_academic_pairs": STATIC_MIN_NON_ACADEMIC_PAIRS,
        "max_weak_evidence_matches": STATIC_MAX_WEAK_EVIDENCE_MATCHES,
        "max_low_score_matches": STATIC_MAX_LOW_SCORE_MATCHES,
        "min_geo_proximity_sum": STATIC_MIN_GEO_PROXIMITY_SUM,
        "max_academic_academic_pairs": STATIC_MAX_ACADEMIC_ACADEMIC_PAIRS,
        "neural_feedback_ranker": neural_summary,
        "min_final_pair_score": MIN_FINAL_PAIR_SCORE,
        "require_implementation_anchor": REQUIRE_IMPLEMENTATION_ANCHOR,
        "candidate_count": len(patterns),
        "selected_count": len(selected_ids),
        "constraints": [
            "exactly MAX_MEETINGS selected matches",
            "each member appears in at most MAX_MEETINGS_PER_MEMBER selected match",
            "each concrete cluster appears in at most MAX_SELECTED_PER_CLUSTER selected matches",
            "each opportunity_id appears in at most MAX_SELECTED_PER_OPPORTUNITY selected matches",
            "role-weighted exposure mass stays within Role_Weights quota limits when present",
            "subrole-weighted exposure mass stays within Subrole_Weights quota limits when present",
            "academic_support/university matches and actor slots stay within SCIP_Portfolio_Constraints caps",
            "at least min_non_academic_pairs selected matches have no academic actor",
            "weak-evidence and low-score matches stay within SCIP_Portfolio_Constraints caps",
            "selected portfolio reaches the workbook geo proximity floor",
            "academic-academic matches stay within the workbook cap",
            "SCIP_Optimization_Pool is preferred as bounded SCIP input when present",
            "Pair_Vector remains the full audit universe and Opportunity_Cards remains the card layer",
        ],
    }
    path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")


def load_json_list(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, list) else []
    except Exception:
        return []


def signal_titles_for_history(raw_signals: Any, limit: int = 5) -> list[str]:
    titles = []
    seen = set()
    for raw in str(raw_signals or "").split("||"):
        raw = clean_display_value(raw)
        if not raw:
            continue
        title = raw.split(":", 1)[1].strip() if ":" in raw else raw
        title = re.split(r"\s{2,}| - ", title, maxsplit=1)[0].strip()
        if len(title) > 120:
            title = title[:117].rstrip() + "..."
        key = title.lower()
        if title and key not in seen:
            seen.add(key)
            titles.append(title)
        if len(titles) >= limit:
            break
    return titles


def load_feedback_label_counts() -> dict[str, Counter]:
    path = PROJECT_DIR / "data" / "feedback" / "scip_feedback.jsonl"
    result: dict[str, Counter] = defaultdict(Counter)
    if not path.exists():
        return result
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            record = json.loads(line)
        except Exception:
            continue
        label = str(record.get("label", "") or record.get("human_feedback_label", "") or "").strip()
        if not label:
            continue
        for key in [record.get("opportunity_id"), record.get("recommendation_id")]:
            if key:
                result[str(key)][label] += 1
    return result


def summarize_feedback_counts(counts: Counter) -> str:
    if not counts:
        return "kein Feedback"
    labels = {
        "useful": "nützlich",
        "interesting_but_weak": "schwach",
        "wrong_connection": "Verbindung falsch",
        "good_topic_wrong_actors": "Akteure falsch",
        "not_relevant": "nicht relevant",
    }
    return ", ".join(f"{count}x {labels.get(label, label)}" for label, count in counts.most_common())



def load_feedback_records() -> list[dict[str, Any]]:
    path = PROJECT_DIR / "data" / "feedback" / "scip_feedback.jsonl"
    records: list[dict[str, Any]] = []
    if not path.exists():
        return records
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            record = json.loads(line)
        except Exception:
            continue
        if isinstance(record, dict):
            records.append(record)
    return records


def feedback_record_label(record: dict[str, Any]) -> str:
    return str(record.get("label", "") or record.get("human_feedback_label", "") or "").strip()


def feedback_record_target(record: dict[str, Any]) -> str:
    return str(record.get("feedback_target", "") or record.get("feedback_dimension", "") or "").strip() or "unspecified"


def feedback_record_reason(record: dict[str, Any]) -> str:
    return str(record.get("reason_category", "") or record.get("feedback_reason_category", "") or "").strip() or "unspecified"


def feedback_selected_value(record: dict[str, Any]) -> Optional[bool]:
    if "selected" not in record:
        return None
    value = record.get("selected")
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "selected"}:
        return True
    if text in {"0", "false", "no", "not_selected"}:
        return False
    return None


HUMAN_FEEDBACK_LABEL_DELTAS = {
    "useful": 7.0,
    "interesting_but_weak": -5.0,
    "wrong_connection": -13.0,
    "good_topic_wrong_actors": -11.0,
    "wrong_actors": -11.0,
    "not_relevant": -16.0,
}

HUMAN_FEEDBACK_REASON_DELTAS = {
    "good_match": 4.0,
    "downgrade_not_drop": 2.0,
    "good_match_weak_evidence": -4.0,
    "good_match_weak_rlp": -5.0,
    "indirect_transfer_legitimacy": -5.0,
    "missing_rlp_application_anchor": -6.0,
    "source_signal_not_application": -7.0,
    "source_only_actor": -7.0,
    "wrong_counterpart": -9.0,
    "insufficient_evidence": -7.0,
    "weak_why_now": -5.0,
    "too_broad_action": -5.0,
    "wrong_pilot_lane": -7.0,
    "too_macro": -8.0,
    "not_public": -10.0,
    "not_publicly_show": -10.0,
}

HUMAN_FEEDBACK_TARGET_DELTAS = {
    "actor_match": 1.0,
    "evidence_signal": -1.0,
    "rlp_relevance": -1.5,
    "next_action": -1.0,
    "typology_stage": -1.0,
    "public_framing": -1.0,
}


def _record_actors(record: dict[str, Any]) -> list[str]:
    actors = record.get("actors")
    if isinstance(actors, list):
        return [str(actor) for actor in actors if str(actor).strip()]
    pair = record.get("pair") or record.get("members")
    if isinstance(pair, str) and pair.strip():
        parts = re.split(r"\s+(?:x|\|)\s+", pair)
        return [part.strip() for part in parts if part.strip()]
    return []


def _human_feedback_relation_weight(record: dict[str, Any], pattern: MeetingPattern) -> tuple[float, str]:
    recommendation_id = str(record.get("recommendation_id", "") or "").strip()
    opportunity_id = str(record.get("opportunity_id", "") or "").strip()
    if recommendation_id and recommendation_id == pattern_recommendation_id(pattern):
        return 1.0, "exact_recommendation"
    if opportunity_id and opportunity_id == opportunity_id_for_pattern(pattern):
        return 0.9, "exact_opportunity"

    record_actors = {normalize_search_text(actor) for actor in _record_actors(record)}
    pattern_actors = {normalize_search_text(actor) for actor in pattern.members}
    actor_overlap = len(record_actors & pattern_actors)
    record_topic = normalize_search_text(
        str(record.get("topic", "") or record.get("convening_theme", "") or record.get("opportunity_key", "") or "")
    )
    pattern_topic = normalize_search_text(pattern.convening_theme)
    topic_match = bool(record_topic and (record_topic in pattern_topic or pattern_topic in record_topic))
    if actor_overlap >= 2 and topic_match:
        return 0.45, "same_actors_topic"
    if actor_overlap >= 2:
        return 0.30, "same_actors"
    if actor_overlap == 1 and topic_match:
        return 0.18, "one_actor_topic"
    if topic_match:
        return 0.12, "same_topic"
    return 0.0, ""


def explicit_human_feedback_adjustment_for_pattern(pattern: MeetingPattern) -> tuple[float, dict[str, Any]]:
    label_total = 0.0
    reason_total = 0.0
    target_total = 0.0
    selection_total = 0.0
    exact_count = 0
    similar_count = 0
    useful_count = 0
    negative_count = 0
    applied_count = 0
    reasons: Counter[str] = Counter()
    labels: Counter[str] = Counter()

    for record in load_feedback_records():
        weight, relation = _human_feedback_relation_weight(record, pattern)
        if weight <= 0:
            continue
        label = feedback_record_label(record)
        reason = feedback_record_reason(record)
        target = feedback_record_target(record)
        selected = feedback_selected_value(record)
        label_delta = HUMAN_FEEDBACK_LABEL_DELTAS.get(label, 0.0)
        reason_delta = HUMAN_FEEDBACK_REASON_DELTAS.get(reason, 0.0)
        target_delta = HUMAN_FEEDBACK_TARGET_DELTAS.get(target, 0.0) if label and label != "useful" else 0.0
        selected_delta = 0.0
        if selected is True and label_delta < 0:
            selected_delta = -2.0
        elif selected is False and label == "useful":
            selected_delta = 3.0
        elif selected is False and label_delta < 0:
            selected_delta = -1.0

        if relation.startswith("exact"):
            exact_count += 1
        else:
            similar_count += 1
        if label == "useful":
            useful_count += 1
        elif label_delta < 0 or reason_delta < 0:
            negative_count += 1
        if label:
            labels[label] += 1
        if reason and reason != "unspecified":
            reasons[reason] += 1

        label_total += weight * label_delta
        reason_total += weight * reason_delta
        target_total += weight * target_delta
        selection_total += weight * selected_delta
        applied_count += 1

    total = label_total + reason_total + target_total + selection_total
    total = max(-30.0, min(18.0, total))
    return total, {
        "human_feedback_count": applied_count,
        "human_feedback_exact_count": exact_count,
        "human_feedback_similar_count": similar_count,
        "human_feedback_positive_count": useful_count,
        "human_feedback_negative_count": negative_count,
        "human_feedback_label_adjustment": label_total,
        "human_feedback_reason_adjustment": reason_total,
        "human_feedback_target_adjustment": target_total,
        "human_feedback_selection_adjustment": selection_total,
        "human_feedback_top_labels": "; ".join(f"{k}:{v}" for k, v in labels.most_common(3)),
        "human_feedback_top_reasons": "; ".join(f"{k}:{v}" for k, v in reasons.most_common(3)),
    }


def feedback_actor_role_pair(record: dict[str, Any]) -> str:
    actor_types = record.get("actor_types")
    roles: list[str] = []
    if isinstance(actor_types, dict):
        roles = [str(value or "").strip() for value in actor_types.values() if str(value or "").strip()]
    if not roles:
        actors = record.get("actors", [])
        if isinstance(actors, list):
            for actor in actors:
                roles.extend(sorted(infer_member_roles(str(actor))))
    if not roles:
        return "unknown"
    normalized_roles = []
    for role in roles:
        role_text = normalize_search_text(role)
        role_text = role_text.replace("profit_anchor", "implementation_actor")
        role_text = role_text.replace("implementation_anchor", "implementation_actor")
        role_text = role_text.replace("academic_support", "knowledge_actor")
        role_text = role_text.replace("context_actor", "visibility_actor")
        normalized_roles.append(role_text)
    return " + ".join(sorted(dict.fromkeys(normalized_roles))[:4])


def summarize_feedback_learning(patterns_df: pd.DataFrame, output_path: Path) -> dict[str, Any]:
    records = load_feedback_records()
    label_counts = Counter()
    target_counts = Counter()
    reason_counts = Counter()
    selected_counts = Counter()
    schema_counts = Counter()
    positive_role_patterns = Counter()
    negative_role_patterns = Counter()
    selected_negative = Counter()
    unselected_positive = Counter()

    positive_labels = {"useful"}
    negative_labels = {"not_relevant", "wrong_connection", "good_topic_wrong_actors"}
    weak_labels = {"interesting_but_weak"}

    current_selected_ids = set()
    current_selected_opportunities = set()
    if patterns_df is not None and not patterns_df.empty:
        selected_df = patterns_df[patterns_df.get("selected", 0).astype(str) == "1"].copy()
        current_selected_ids = {
            clean_display_value(value)
            for value in selected_df.get("recommendation_id", [])
            if clean_display_value(value)
        }
        current_selected_opportunities = {
            clean_display_value(value)
            for value in selected_df.get("opportunity_id", [])
            if clean_display_value(value)
        }

    for record in records:
        label = feedback_record_label(record)
        if not label:
            continue
        label_counts[label] += 1
        target_counts[feedback_record_target(record)] += 1
        reason_counts[feedback_record_reason(record)] += 1

        if record.get("recommendation_id"):
            schema_counts["has_recommendation_id"] += 1
        if record.get("opportunity_id"):
            schema_counts["has_opportunity_id"] += 1
        if "selected" in record:
            schema_counts["has_selected"] += 1
        if feedback_record_target(record) != "unspecified":
            schema_counts["has_feedback_target"] += 1
        if feedback_record_reason(record) != "unspecified":
            schema_counts["has_reason_category"] += 1

        selected_value = feedback_selected_value(record)
        if selected_value is None:
            recommendation_id = clean_display_value(record.get("recommendation_id", ""))
            opportunity_id = clean_display_value(record.get("opportunity_id", ""))
            if recommendation_id in current_selected_ids or opportunity_id in current_selected_opportunities:
                selected_value = True
        if selected_value is None:
            selected_counts["unknown"] += 1
        elif selected_value:
            selected_counts["selected"] += 1
        else:
            selected_counts["not_selected"] += 1

        role_pair = feedback_actor_role_pair(record)
        if label in positive_labels:
            positive_role_patterns[role_pair] += 1
            if selected_value is False:
                unselected_positive[role_pair] += 1
        elif label in negative_labels:
            negative_role_patterns[role_pair] += 1
            if selected_value is True:
                selected_negative[role_pair] += 1
        elif label in weak_labels:
            negative_role_patterns[f"weak: {role_pair}"] += 1

    total = sum(label_counts.values())
    summary = {
        "feedback_items_loaded": total,
        "label_counts": dict(label_counts.most_common()),
        "target_counts": dict(target_counts.most_common()),
        "reason_counts": dict(reason_counts.most_common()),
        "selected_status_counts": dict(selected_counts.most_common()),
        "schema_coverage": {
            "recommendation_id": schema_counts["has_recommendation_id"],
            "opportunity_id": schema_counts["has_opportunity_id"],
            "selected": schema_counts["has_selected"],
            "feedback_target": schema_counts["has_feedback_target"],
            "reason_category": schema_counts["has_reason_category"],
        },
        "strongest_positive_patterns": dict(positive_role_patterns.most_common(6)),
        "strongest_negative_patterns": dict(negative_role_patterns.most_common(6)),
        "selected_but_negative_patterns": dict(selected_negative.most_common(6)),
        "not_selected_but_useful_patterns": dict(unselected_positive.most_common(6)),
        "most_common_issue": reason_counts.most_common(1)[0][0] if reason_counts else "",
    }
    output_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Feedback learning summary:")
    print(f" - {total} feedback items loaded")
    if label_counts:
        print(" - labels: " + ", ".join(f"{key}={value}" for key, value in label_counts.most_common()))
    if target_counts:
        print(" - targets: " + ", ".join(f"{key}={value}" for key, value in target_counts.most_common(6)))
    if reason_counts:
        print(" - most common issue: " + reason_counts.most_common(1)[0][0])
    if positive_role_patterns:
        key, value = positive_role_patterns.most_common(1)[0]
        print(f" - strongest positive pattern: {key} ({value})")
    if negative_role_patterns:
        key, value = negative_role_patterns.most_common(1)[0]
        print(f" - strongest negative pattern: {key} ({value})")
    print(f" - summary file: {output_path}")
    return summary



def collapse_opportunity_history(history: list[dict[str, Any]]) -> list[dict[str, Any]]:
    collapsed: dict[tuple[str, str, str], dict[str, Any]] = {}
    for item in sorted(history, key=lambda row: str(row.get("created_at", ""))):
        opportunity_id = str(item.get("opportunity_id", "") or "").strip()
        date = str(item.get("date", "") or "").strip()
        source = str(item.get("source", "") or "").strip()
        if not opportunity_id or not date:
            continue
        collapsed[(opportunity_id, date, source)] = item
    return sorted(collapsed.values(), key=lambda item: str(item.get("created_at", "")))


def update_opportunity_history(patterns_df: pd.DataFrame, *, source_name: str, timestamp: str) -> None:
    if patterns_df.empty or "opportunity_id" not in patterns_df.columns:
        return
    history = collapse_opportunity_history(load_json_list(OPPORTUNITY_HISTORY_PATH))
    today = datetime.now().strftime("%Y-%m-%d")
    feedback_counts = load_feedback_label_counts()
    new_records = []
    replace_keys: set[tuple[str, str, str]] = set()
    selected = patterns_df[patterns_df.get("selected", 0).astype(str) == "1"].copy()
    if selected.empty:
        return
    for _, row in selected.iterrows():
        opportunity_id = clean_display_value(row.get("opportunity_id", ""))
        if not opportunity_id:
            continue
        key = (opportunity_id, today, source_name)
        replace_keys.add(key)
        prior = [
            item
            for item in history
            if str(item.get("opportunity_id")) == opportunity_id
            and (str(item.get("date")), str(item.get("source"))) != (today, source_name)
        ]
        prior_points = [int(item.get("score_points", 0) or 0) for item in prior]
        points = int(row.get("display_points", 0) or scip_display_points(row))
        if not prior_points:
            trend = "neu"
            first_seen = datetime.now().strftime("%Y-%m-%d")
            previous_status = ""
        else:
            last_points = prior_points[-1]
            trend = "steigend" if points > last_points else "fallend" if points < last_points else "stabil"
            first_seen = str(prior[0].get("first_seen") or prior[0].get("date") or datetime.now().strftime("%Y-%m-%d"))
            previous_status = str(prior[-1].get("status", ""))
        signals = signal_titles_for_history(row.get("recent_signals", ""))
        old_signals = {signal for item in prior for signal in item.get("signals", [])}
        feedback = feedback_counts.get(opportunity_id, Counter()) + feedback_counts.get(str(row.get("recommendation_id", "")), Counter())
        new_records.append({
            "date": today,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "run_id": timestamp,
            "source": source_name,
            "opportunity_id": opportunity_id,
            "recommendation_id": clean_display_value(row.get("recommendation_id", "")),
            "actors": [part.strip() for part in clean_display_value(row.get("members", "")).split("|") if part.strip()],
            "topic": clean_display_value(row.get("convening_theme", "")),
            "status": clean_display_value(row.get("public_status", "")) or public_status_for_points(points),
            "previous_status": previous_status,
            "score": float(row.get("score", 0) or 0),
            "score_points": points,
            "previous_score_points": prior_points[-1] if prior_points else None,
            "trend": trend,
            "first_seen": first_seen,
            "new_signal_count": len([signal for signal in signals if signal not in old_signals]),
            "signals": signals,
            "feedback_summary": summarize_feedback_counts(feedback),
        })
    if not new_records and not replace_keys:
        return
    history = [
        item for item in history
        if (str(item.get("opportunity_id")), str(item.get("date")), str(item.get("source"))) not in replace_keys
    ]
    history = collapse_opportunity_history(history + new_records)[-750:]
    OPPORTUNITY_HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    OPPORTUNITY_HISTORY_PATH.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Opportunity History: {len(new_records)} Entwicklungslinien aktualisiert -> {OPPORTUNITY_HISTORY_PATH}")


def classify_convening(
        score: float,
        evidence_count: int,
        rlp_fit: str,
        common_problem_quality: str,
        distinct_evidence_members: int = 0,
        non_generic_question: bool = True,
        pair_balance: str = "balanced",
        pair_balance_exception: bool = False,
) -> dict[str, Any]:
    signal_strength = round(score)
    balanced_or_exception = pair_balance == "balanced" or bool(pair_balance_exception)
    if pair_balance == "macro_only" and score >= 25:
        return {
            "zirp_fit": "Unklar",
            "aktionsreife": "Niedrig",
            "signalstaerke": signal_strength,
            "empfehlung": "WATCH / Briefing",
        }
    if (
        score >= 45
        and evidence_count >= 2
        and distinct_evidence_members >= 2
        and rlp_fit in ["hoch", "mittel"]
        and common_problem_quality in ["stark", "mittel"]
        and non_generic_question
        and balanced_or_exception
    ):
        return {
            "zirp_fit": "Hoch",
            "aktionsreife": "Mittel",
            "signalstaerke": signal_strength,
            "empfehlung": "DO / Sondieren",
        }
    if score >= 35 and evidence_count >= 2 and common_problem_quality in ["mittel", "stark"]:
        return {
            "zirp_fit": "Mittel",
            "aktionsreife": "Niedrig bis Mittel",
            "signalstaerke": signal_strength,
            "empfehlung": "WATCH / Beobachten",
        }
    if score >= 25:
        return {
            "zirp_fit": "Unklar",
            "aktionsreife": "Niedrig",
            "signalstaerke": signal_strength,
            "empfehlung": "MAYBE / Reframen",
        }
    return {
        "zirp_fit": "Niedrig",
        "aktionsreife": "Niedrig",
        "signalstaerke": signal_strength,
        "empfehlung": "DROP / Nicht weiterverfolgen",
    }



def extract_json_object(text: str) -> dict[str, Any]:
    if not text:
        return {}
    cleaned = re.sub(r"^```(?:json)?", "", text.strip(), flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return {}
    try:
        parsed = json.loads(cleaned[start:end + 1])
    except Exception:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def load_convening_report_context() -> str:
    context_path = os.getenv("ZIRP_CONVENING_CONTEXT_FILE", "").strip()
    if not context_path:
        return ""
    path = Path(context_path)
    if not path.exists():
        return ""
    try:
        text = path.read_text(encoding="utf-8")
    except Exception:
        return ""
    return re.sub(r"\s+", " ", text).strip()[:4500]



def selected_patterns_df(patterns_df: pd.DataFrame) -> pd.DataFrame:
    if patterns_df.empty or "selected" not in patterns_df.columns:
        return pd.DataFrame()
    return patterns_df[patterns_df["selected"].astype(int) == 1].copy()


def split_pattern_members(value: Any) -> list[str]:
    return [member.strip() for member in str(value or "").split("|") if member.strip()]


REQUIRED_PATTERN_ROLES = {"problem_owner", "implementation_actor", "knowledge_actor"}


def actor_use_mode_for_member(member_name: str) -> str:
    key = normalize_search_text(member_name)
    if key in BROAD_MEDIA_MEMBERS_NORM or any(marker in key for marker in ["swr", "rundfunk", "zeitung", "verlag", "media"]):
        return "visibility_actor"
    if any(marker in key for marker in ["universitat", "hochschule", "institut", "forschung"]):
        return "knowledge_actor"
    if any(marker in key for marker in ["bank", "sparkasse", "isb", "versicherung", "ministerium"]):
        return "resource_actor"
    if any(marker in key for marker in ["ministerium", "stadt", "kreis", "agentur", "kammer", "verband", "aok", "ikk"]):
        return "problem_owner"
    if is_implementation_anchor_member(member_name):
        return "implementation_actor"
    return "relationship_broker"


def infer_member_roles(member_name: str) -> set[str]:
    role = actor_role(member_name)
    key = normalize_search_text(member_name)
    roles: set[str] = set()
    if role in {"profit_anchor", "implementation_anchor"}:
        roles.add("implementation_actor")
    if role == "academic_support":
        roles.add("knowledge_actor")
    if role == "context_actor":
        roles.add("public_interest_voice")
    if any(marker in key for marker in ["iqib", "zahnen", "technik", "werke", "gmbh", "industrie", "bau", "energie", "wasser"]):
        roles.add("implementation_actor")
    if any(marker in key for marker in ["ministerium", "stadt", "kreis", "agentur", "kammer", "verband", "krankenkasse", "aok", "ikk", "bundesagentur"]):
        roles.add("problem_owner")
    if any(marker in key for marker in ["bank", "sparkasse", "isb", "versicherung", "ministerium", "lbbw", "landesbank"]):
        roles.add("resource_actor")
        roles.add("resource_holder")
        # Financial/resource actors can unlock implementation, but they are not themselves the application context.
        if any(marker in key for marker in ["bank", "sparkasse", "versicherung", "lbbw", "landesbank"]):
            roles.discard("implementation_actor")
    if any(marker in key for marker in ["ministerium", "vorstand", "geschaeftsfuehrung", "leitung", "stadt", "kreis"]):
        roles.add("decision_maker")
    if any(marker in key for marker in ["universitat", "university", "hochschule", "institut", "forschung", "iqib"]):
        roles.add("knowledge_actor")
        roles.add("technical_or_domain_expert")
        roles.add("content_expert")
    if key in BROAD_MEDIA_MEMBERS_NORM or any(marker in key for marker in ["swr", "rundfunk", "zeitung", "verlag", "media"]):
        roles.add("visibility_actor")
        roles.add("media_actor")
    if any(marker in key for marker in ["zirp", "kammer", "verband", "netzwerk"]):
        roles.add("relationship_broker")
        roles.add("coalition_anchor")
    return roles


def missing_roles_for_typology(member_roles: set[str], typology: ConveningTypology | str) -> set[str]:
    try:
        typology_value = typology if isinstance(typology, ConveningTypology) else ConveningTypology(str(typology))
    except ValueError:
        return set()
    return set(TYPOLOGY_REQUIRED_ROLES.get(typology_value, set())) - set(member_roles)


def audit_institutional_dominance(patterns_df: pd.DataFrame, output_path: Path) -> pd.DataFrame:
    selected = selected_patterns_df(patterns_df)
    rows = []

    if selected.empty:
        audit_df = pd.DataFrame([{
            "metric": "selected_patterns",
            "value": 0,
            "status": "OK",
            "note": "No selected patterns.",
        }])
        audit_df.to_csv(output_path, index=False, encoding="utf-8-sig")
        return audit_df

    selected_count = len(selected)
    all_member_slots = []
    high_gravity_slots = []
    public_system_slots = []
    media_slots = []
    selected_with_high_gravity = 0
    selected_with_public_system = 0
    selected_with_media = 0
    selected_do_count = 0
    selected_watch_count = 0
    selected_maybe_count = 0
    selected_drop_count = 0
    total_bonus = 0.0
    bonus_count = 0

    for _, row in selected.iterrows():
        members = split_pattern_members(row.get("members", ""))
        all_member_slots.extend(members)
        row_high = [member for member in members if member in HIGH_GRAVITY_MEMBERS]
        row_public = [member for member in members if member in PUBLIC_SYSTEM_MEMBERS]
        row_media = [member for member in members if member in BROAD_MEDIA_MEMBERS_AUDIT]
        high_gravity_slots.extend(row_high)
        public_system_slots.extend(row_public)
        media_slots.extend(row_media)
        selected_with_high_gravity += int(bool(row_high))
        selected_with_public_system += int(bool(row_public))
        selected_with_media += int(bool(row_media))

        empfehlung = str(row.get("empfehlung", "")).lower()
        if "do" in empfehlung:
            selected_do_count += 1
        elif "watch" in empfehlung:
            selected_watch_count += 1
        elif "maybe" in empfehlung:
            selected_maybe_count += 1
        elif "drop" in empfehlung:
            selected_drop_count += 1

        structural_value = float(row.get("structural_bonus", 0) or 0)
        if structural_value > 0:
            total_bonus += structural_value
            bonus_count += 1

        roles = set()
        for member in members:
            roles |= infer_member_roles(member)
        missing_roles = REQUIRED_PATTERN_ROLES - roles
        if missing_roles:
            rows.append({
                "metric": "missing_required_roles",
                "value": len(missing_roles),
                "status": "WARN",
                "note": f"{row.get('pattern_id', '')}: {', '.join(sorted(missing_roles))}",
            })

    total_slots = len(all_member_slots) or 1
    metrics = {
        "high_gravity_slot_share": len(high_gravity_slots) / total_slots,
        "public_system_slot_share": len(public_system_slots) / total_slots,
        "media_slot_share": len(media_slots) / total_slots,
        "pattern_high_gravity_share": selected_with_high_gravity / selected_count,
        "pattern_public_system_share": selected_with_public_system / selected_count,
        "pattern_media_share": selected_with_media / selected_count,
        "avg_structural_bonus": total_bonus / bonus_count if bonus_count else 0.0,
    }

    def status_for(metric: str, value: float) -> str:
        thresholds = {
            "high_gravity_slot_share": 0.55,
            "public_system_slot_share": 0.50,
            "media_slot_share": 0.35,
            "pattern_high_gravity_share": 0.80,
            "pattern_public_system_share": 0.75,
            "pattern_media_share": 0.50,
            "avg_structural_bonus": 4.0,
        }
        limit = thresholds.get(metric)
        return "WARN" if limit is not None and value > limit else "OK"

    rows.append({"metric": "selected_patterns", "value": selected_count, "status": "OK", "note": "Number of selected Convening Radar patterns."})
    for metric, value in metrics.items():
        rows.append({"metric": metric, "value": round(value, 3), "status": status_for(metric, value), "note": "Dominance/structural scoring audit metric."})
    rows.extend([
        {"metric": "do_count", "value": selected_do_count, "status": "OK", "note": "Selected patterns marked DO."},
        {"metric": "watch_count", "value": selected_watch_count, "status": "OK", "note": "Selected patterns marked WATCH."},
        {"metric": "maybe_count", "value": selected_maybe_count, "status": "OK", "note": "Selected patterns marked MAYBE."},
        {"metric": "drop_count", "value": selected_drop_count, "status": "WARN" if selected_drop_count else "OK", "note": "Selected patterns marked DROP should normally be zero."},
    ])

    for member, count in Counter(all_member_slots).most_common(20):
        rows.append({"metric": "top_selected_member", "value": count, "status": "WARN" if count >= 3 else "OK", "note": member})

    audit_df = pd.DataFrame(rows)
    audit_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    warnings = audit_df[audit_df["status"] == "WARN"]
    if not warnings.empty:
        print("Dominance audit warnings:")
        for _, row in warnings.iterrows():
            print(f" - {row['metric']}: {row['value']} ({row['note']})")
    return audit_df


def apply_convening_selection_filters(patterns_df: pd.DataFrame) -> pd.DataFrame:
    if patterns_df.empty:
        return patterns_df

    df = patterns_df.copy()
    selected_mask = df.get("selected", 0).astype(int) == 1

    format_text = df.get("suggested_format", "").astype(str).str.lower()
    empfehlung_text = df.get("empfehlung", "").astype(str).str.lower()
    drop_mask = (
        format_text.str.contains("nicht weiterverfolgen", na=False)
        | empfehlung_text.str.contains("drop", na=False)
        | empfehlung_text.str.contains("nicht weiterverfolgen", na=False)
    )

    problem_mask = df.apply(
        lambda row: has_clear_convening_problem_text(" ".join([
            str(row.get("convening_theme", "")),
            str(row.get("why_these_members", "")),
            str(row.get("editorial_justification", "")),
            str(row.get("possible_agenda", "")),
            str(row.get("recent_signals", "")),
        ])),
        axis=1,
    )

    static_source_mask = df.get("candidate_source", "").astype(str).eq("static_excel")

    df["editorial_gate_status"] = "pass"
    df["editorial_gate_reason"] = ""
    df.loc[selected_mask & drop_mask & ~static_source_mask, "editorial_gate_status"] = "fail_drop"
    df.loc[selected_mask & drop_mask & ~static_source_mask, "editorial_gate_reason"] = "Drop/Nicht-weiterverfolgen recommendation"
    df.loc[selected_mask & drop_mask & static_source_mask, "editorial_gate_status"] = "static_selected_low_display_score"
    df.loc[selected_mask & drop_mask & static_source_mask, "editorial_gate_reason"] = "Static Excel candidate selected by SCIP although display classifier marks it weak"
    unclear_mask = selected_mask & ~problem_mask
    df.loc[unclear_mask & ~static_source_mask, "editorial_gate_status"] = "fail_unclear_problem"
    df.loc[unclear_mask & ~static_source_mask, "editorial_gate_reason"] = "No concrete problem plus action verb"
    df.loc[unclear_mask & static_source_mask & ~drop_mask, "editorial_gate_status"] = "static_needs_live_validation"
    df.loc[unclear_mask & static_source_mask & ~drop_mask, "editorial_gate_reason"] = "Static Excel candidate selected by SCIP; live problem/action evidence should be validated"

    remove_mask = selected_mask & ~static_source_mask & drop_mask
    generated_unclear_remove_mask = selected_mask & ~static_source_mask & ~problem_mask
    remove_mask = remove_mask | generated_unclear_remove_mask
    if remove_mask.any():
        df.loc[remove_mask, "selected"] = 0
        removed = int(remove_mask.sum())
        print(f"Convening editorial gates: {removed} ausgewÃ¤hlte VorschlÃ¤ge entfernt")

    return df


def convening_model_payload(patterns_df: pd.DataFrame, report_context: str = "") -> list[dict[str, Any]]:
    selected = patterns_df[patterns_df["selected"] == 1].copy()
    if selected.empty:
        return []
    selected["score_num"] = pd.to_numeric(selected["score"], errors="coerce").fillna(0)
    selected = selected.sort_values("score_num", ascending=False).head(MAX_MODEL_CONVENING_CARDS)
    contrast_df = patterns_df.copy()
    contrast_df["score_num"] = pd.to_numeric(contrast_df.get("score", 0), errors="coerce").fillna(0)
    contrast_df = contrast_df.sort_values("score_num", ascending=False).head(20)
    contrast_candidates = [
        {
            "pattern_id": int(other.get("pattern_id", 0)),
            "members": str(other.get("members", "")),
            "score": str(other.get("score", "")),
            "empfehlung": str(other.get("empfehlung", "")),
            "convening_typology": str(other.get("convening_typology", "")),
            "convening_stage": str(other.get("convening_stage", "")),
            "purpose_maturity": str(other.get("purpose_maturity", "")),
            "next_engagement_move": str(other.get("next_engagement_move", "")),
            "reason": str(other.get("editorial_justification", "")),
        }
        for _, other in contrast_df.iterrows()
    ]
    payload = []
    for position, (_, row) in enumerate(selected.iterrows()):
        card_label = scip_card_label(row, position)
        payload.append({
            "pattern_id": int(row.get("pattern_id", 0)),
            "members": str(row.get("members", "")),
            "actors": str(row.get("members", "")),
            "score": int(scip_display_points(row)),
            "score_scale": 100,
            "card_label": card_label,
            "card_title": card_title_text(row, card_label),
            "python_decision": str(row.get("action_decision", "")),
            "empfehlung": str(row.get("empfehlung", "")),
            "zirp_fit": str(row.get("zirp_fit", "")),
            "aktionsreife": str(row.get("aktionsreife", "")),
            "signalstaerke": str(row.get("signalstaerke", "")),
            "convening_theme": str(row.get("convening_theme", "")),
            "convening_stage": str(row.get("convening_stage", "")),
            "purpose_maturity": str(row.get("purpose_maturity", "")),
            "next_engagement_move": str(row.get("next_engagement_move", "")),
            "do_not_ask_for": str(row.get("do_not_ask_for", "")),
            "north_star_purpose": str(row.get("north_star_purpose", "")),
            "participants_logic": str(row.get("participants_logic", "")),
            "required_inputs": str(row.get("required_inputs", "")),
            "expected_outputs": str(row.get("expected_outputs", "")),
            "next_decision_gate": str(row.get("next_decision_gate", "")),
            "rlp_relevance_type": rlp_relevance_type(row),
            "pair_evidence_balance": str(row.get("pair_evidence_balance", "")),
            "pair_evidence_balance_exception": str(row.get("pair_evidence_balance_exception", "")),
            "evidence_quality": str(row.get("evidence_quality", "")),
            "actor_use_modes": str(row.get("actor_use_modes", "")),
            "match_logic": str(row.get("editorial_justification", "")),
            "evidence_reading": str(row.get("why_now", "")),
            "main_weakness": str(row.get("uncertainty", "")),
            "missing_anchor": str(row.get("missing_typology_roles", "")),
            "recommended_next_step": str(row.get("next_best_action", "")),
            "do_not_assume": str(row.get("do_not_ask_for", "")),
            "shared_topics": str(row.get("shared_topics", "")),
            "bridge_topics": str(row.get("bridge_topics", "")),
            "concrete_clusters": str(row.get("concrete_clusters", "")),
            "cluster_problem": str(row.get("cluster_problem", "")),
            "python_reason": str(row.get("editorial_justification", "")),
            "signals": [
                signal.strip()
                for signal in str(row.get("recent_signals", "")).split("||")
                if signal.strip()
            ][:4],
            "suggested_format": str(row.get("suggested_format", "")),
            "possible_agenda": str(row.get("possible_agenda", "")),
        })
    return payload


def validate_model_enrichment(parsed: dict[str, Any], source_item: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    source_members = {
        member.strip()
        for member in str(source_item.get("members", "")).split("|")
        if member.strip()
    }

    actors = parsed.get("actors", [])
    if isinstance(actors, list):
        unknown_actors = [str(actor).strip() for actor in actors if str(actor).strip() and str(actor).strip() not in source_members]
        if unknown_actors:
            errors.append(f"Unknown actors introduced: {unknown_actors}")

    next_step = normalize_search_text(parsed.get("recommended_next_step", ""))
    if not next_step:
        errors.append("Missing recommended_next_step")

    broad_terms = ["mehrere sektoren", "verschiedene bereiche", "breit diskutieren", "stakeholderdialog"]
    if any(term in next_step for term in broad_terms):
        errors.append("Recommended next step is too broad")

    evidence = parsed.get("evidence_signals", [])
    if isinstance(evidence, list) and len([item for item in evidence if str(item).strip()]) == 0:
        errors.append("No evidence signals cited")

    reason_1 = normalize_search_text(parsed.get("complementarity_reason_1", ""))
    reason_2 = normalize_search_text(parsed.get("complementarity_reason_2", ""))
    if not reason_1 or not reason_2:
        errors.append("Missing two explicit complementarity reasons")
    elif reason_1 == reason_2:
        errors.append("Complementarity reasons are redundant")

    return errors


def enrich_selected_patterns_with_ollama(patterns_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    report_context = load_convening_report_context()
    payload = convening_model_payload(patterns_df, report_context=report_context)
    if not payload:
        return patterns_df, 0

    client = build_openai_compatible_client()
    if client is None:
        return patterns_df, 0

    enriched = patterns_df.copy()
    model = (os.getenv("ZIRP_CONVENING_MODEL", "").strip() or os.getenv("OLLAMA_CONVENING_MODEL", "").strip() or "qwen2.5:7b-instruct")
    system_prompt = (
        "You are writing short SCIP Convening Radar cards for ZIRP. "
        "The Python/SCIP fields are binding. Do not change the score, selected status, recommendation, convening stage, "
        "RLP relevance, actor roles, evidence quality, or next action. Your task is only to write the visible card text in clear German. "
        "Return valid JSON only."
    )
    legacy_strategy_prompt = (
        "Du bist SCIP Convening Strategist fuer ZIRP. Du erhaeltst angereicherte News-Signale mit Opportunity-Substanz "
        "und leitest daraus sinnvolle SCIP-Matchings und Convening-Kandidaten ab. Kombiniere Akteure nicht wegen aehnlicher Keywords, "
        "sondern nur bei echter strategischer Komplementaritaet: Bedarf trifft Expertise, Tool trifft Anwendungskontext, "
        "wissenschaftliche Kompetenz trifft Praxiszugang, Reichweite trifft Umsetzungskapazitaet, unterschiedliche Perspektiven "
        "auf dieselbe regionale Herausforderung, aktuelles Zeitfenster oder realistische Chance auf Gespraech, Pilot oder Transferformat. "
        "Formuliere die Komplementaritaet paar-spezifisch: Welche konkrete Faehigkeit bringt Akteur A, welchen Anwendungskontext "
        "oder Bedarf bringt Akteur B, und welches testbare Ergebnis koennte daraus entstehen? Vermeide breite Saetze wie "
        "'Regulatorischer Druck trifft auf Investitions- und Umsetzungsfragen'. Besser: 'Digitale Anlagen- und Prozesskompetenz "
        "trifft auf anwendungsnahen Transfer' oder 'Kommunale Resilienz trifft auf ethische und gesellschaftliche Legitimation digitaler Loesungen'. "
        "Medienakteure wie SWR sind nur dann Convening-Akteure, wenn sie mehr als Quelle oder Sichtbarkeit liefern; sonst als Kontextakteur "
        "kennzeichnen oder niedriger priorisieren. "
        "Gib ausschliesslich valides JSON zurueck."
    )
    updated = 0

    for item in payload:
        pattern_id = int(item.get("pattern_id", -1))
        user_prompt = (
            "Write one compact card using exactly this structure:\n"
            "[ACTORS] · [SCORE]/100 · [CARD_LABEL]\n\n"
            "[DECISION_SUMMARY]\n\n"
            "Decision: [DECISION_LINE]\n"
            "Risk: [RISK_LINE]\n"
            "Next: [NEXT_LINE]\n\n"
            "Rules:\n"
            "- DECISION_SUMMARY must be fluent German prose, not bullet points.\n"
            "- DECISION_SUMMARY must be max 4 lines.\n"
            "- It must explain why the actors fit, what the evidence actually shows, what is weak or uncertain, and what should happen next.\n"
            "- Be concrete and critical. Do not make weak signals sound stronger than they are.\n"
            "- If RLP relevance is indirect, say this clearly.\n"
            "- If pair_evidence_balance is one_sided, say: Die Akteurslogik ist plausibel, aber der aktuelle Anlass kommt vor allem von [actor].\n"
            "- If pair_evidence_balance is macro_only, say: Die Signale stuetzen eher ein Briefing als ein Convening.\n"
            "- If a signal is international or outside Rheinland-Pfalz, say that it does not yet prove direct RLP application.\n"
            "- If an actor is only a visibility/source/context actor, say that it is not yet an implementation partner.\n"
            "- If the recommended action is brief, do not call it a convening.\n"
            "- If the recommended action is watch, do not call it a recommendation to act.\n"
            "- If the recommended action is sondieren, describe a validation or sounding step, not a public event.\n"
            "- Do not invent new actors, signals, projects, or pilot fields.\n"
            "- Do not use generic hypotheses such as regulatorischer Druck trifft auf konkrete Pilotfaehigkeit unless the input explicitly contains regulation and company pilot evidence.\n"
            "- If the pair is Bundesagentur x university/hochschule, prefer Arbeitsmarkt, Qualifizierung, Weiterbildung, Fachkraefte, or Bildungstransfer; do not frame it as a digital company pilot unless the evidence explicitly says so.\n"
            "- Do not say Top-Match unless Python provided that label.\n\n"
            "Return valid JSON only with this schema:\n"
            "{\"decision_summary\":\"\",\"decision_line\":\"\",\"risk_line\":\"\",\"next_line\":\"\"}\n\n"
            "Python already owns card_title, score, label and selected status. Do not rewrite the title.\n\n"
            f"INPUT:\n{json.dumps(item, ensure_ascii=False)}"
        )
        def request_card_json(prompt: str) -> tuple[dict[str, Any], str]:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.25,
                max_tokens=OLLAMA_CONVENING_MAX_TOKENS,
                extra_body={"options": {"num_predict": OLLAMA_CONVENING_MAX_TOKENS, "temperature": 0.25, "top_p": 0.9, "repeat_penalty": 1.1}},
            )
            content = response.choices[0].message.content if response.choices else ""
            return extract_json_object(content or "") or {}, content or ""

        try:
            parsed, raw_text = request_card_json(user_prompt)
            required = ["decision_summary", "decision_line", "risk_line", "next_line"]
            if not all(str(parsed.get(key, "") or "").strip() for key in required):
                repair_prompt = (
                    "The previous response was invalid or incomplete. Return ONLY one JSON object with exactly these keys: "
                    "decision_summary, decision_line, risk_line, next_line. Use concise German. No markdown.\n\n"
                    f"INPUT:\n{json.dumps(item, ensure_ascii=False)}\n\n"
                    f"PREVIOUS_RESPONSE:\n{raw_text[:1200]}"
                )
                parsed, _ = request_card_json(repair_prompt)
        except Exception as exc:
            print(f"KI Convening Text: Karte {pattern_id} uebersprungen ({exc})")
            continue

        required = ["decision_summary", "decision_line", "risk_line", "next_line"]
        if not all(str(parsed.get(key, "") or "").strip() for key in required):
            print(f"KI Convening Text: Karte {pattern_id} ohne gueltiges JSON; nutze Python-Fallback")
            continue

        item_members_text = normalize_search_text(item.get("members", ""))
        item_balance = normalize_search_text(item.get("pair_evidence_balance", ""))
        item_decision = normalize_search_text(item.get("python_decision", ""))
        is_macro_or_brief = item_balance == "macro_only" or item_decision == "brief"
        is_bundesagentur_education = (
            "bundesagentur" in item_members_text
            and any(marker in item_members_text for marker in ["universitat", "universitaet", "uni trier", "hochschule"])
        )

        if is_macro_or_brief:
            parsed["decision_line"] = "Als Briefinglinie behalten; nicht als Convening oder Sondierung ausgeben."
            parsed["risk_line"] = "Die Signale stuetzen eher ein Briefing als ein Convening und belegen noch keinen konkreten RLP-Umsetzungsfall."
            parsed["next_line"] = "Internes Briefing oder Beobachtungslinie formulieren; erst bei konkretem Anwendungspartner sondieren."
        elif item_decision == "watch" or item_balance in {"one_sided", "indirect_only"}:
            parsed["decision_line"] = "Beobachten und fehlenden Evidenzanker validieren; noch keine Handlungsempfehlung ableiten."
            if item_balance == "one_sided":
                parsed["risk_line"] = "Die Akteurslogik ist plausibel, aber der aktuelle Anlass kommt vor allem von einer Seite."
            elif item_balance == "indirect_only":
                parsed["risk_line"] = "Die Verbindung ist indirekt und braucht einen konkreten RLP-Anwendungsanker."
            parsed["next_line"] = "Zweiten aktuellen Bedarf oder konkrete RLP-Anwendung pruefen; noch kein Format ansetzen."

        if is_bundesagentur_education and not is_macro_or_brief:
            parsed["decision_summary"] = (
                "Die Bundesagentur bringt einen regionalen Arbeitsmarkt- und Qualifizierungsdruck ein, "
                "waehrend der Bildungsakteur Transfer-, Weiterbildungs- oder Reflexionskompetenz beisteuert. "
                "Die Verbindung ist plausibel, wenn sie auf Fachkraefte, Weiterbildung, Arbeitsmarktintegration "
                "oder Bildungstransfer in Rheinland-Pfalz zugespitzt wird. Vor einem Treffen muss ZIRP klaeren, "
                "welche konkrete Qualifizierungs- oder Arbeitsmarktfrage beide Seiten gemeinsam bearbeiten koennen."
            )
            parsed["decision_line"] = "Reframen auf Arbeitsmarkt, Qualifizierung und Bildungstransfer vor einer Sondierung."
            parsed["risk_line"] = "Der aktuelle Anlass belegt noch keinen digitalen Unternehmenspilot."
            parsed["next_line"] = "Vorkontakt zur Klaerung einer konkreten RLP-Qualifizierungsfrage."

        mask = enriched["pattern_id"].astype(int) == pattern_id
        python_title = str(item.get("card_title", "") or "").strip()
        if python_title and "card_title" in enriched.columns:
            enriched.loc[mask, "card_title"] = python_title

        field_map = {
            "decision_summary": "decision_summary",
            "decision_line": "decision_line",
            "risk_line": "risk_line",
            "next_line": "next_line",
        }
        changed = False
        for source, target in field_map.items():
            value = parsed.get(source, "")
            if isinstance(value, list):
                value = ", ".join(str(v) for v in value if str(v).strip())
            value = repair_mojibake(str(value or "").strip())
            cleanup_replacements = {
                "Hochschule Trier": "Universität Trier",
                "Hochschule Uni Trier": "Universität Trier",
                "Umsetzungspossibiliten": "Umsetzungsmöglichkeiten",
                "Umsetzungspossibilitäten": "Umsetzungsmöglichkeiten",
                "Bilateraler Sondierungsgespräch": "Bilaterales Sondierungsgespräch",
                "bilateraler Sondierungsgespräch": "bilaterales Sondierungsgespräch",
                "Die aktuelle Aktivitäten": "Die aktuellen Aktivitäten",
            }
            for bad, good in cleanup_replacements.items():
                value = value.replace(bad, good)
            if value and target in enriched.columns:
                enriched.loc[mask, target] = value
                changed = True
        if changed:
            updated += 1

    if updated:
        print(f"KI Convening Text: {updated} kurze Karten geschaerft")
    else:
        print("KI Convening Text: keine Karten geschaerft")
    return enriched, updated



def short_signal_title(title: str) -> str:
    lowered = title.lower()
    if "einbaukarte" in lowered or "ersatzbaustoff" in lowered:
        return "EBV-Einbaukarte"
    if "hannover messe" in lowered:
        return "Hannover Messe 2026"
    if "mainzed" in lowered:
        return "mainzed"
    if len(title) > 80:
        return title[:77].rstrip() + "..."
    return title


def short_signal_items(raw_signals: str) -> list[tuple[str, str]]:
    items = []
    for raw in repair_mojibake(str(raw_signals or "")).split("||"):
        raw = raw.strip()
        if not raw:
            continue
        if ":" in raw:
            member, rest = raw.split(":", 1)
        else:
            member, rest = "", raw
        title = re.split(r"\s+[\-]\s+", rest, maxsplit=1)[0].strip()
        items.append((member.strip(), short_signal_title(title)))
    return items


def compact_member_name(member: str) -> str:
    value = clean_display_value(member)
    lower = value.lower()
    if "iqib" in lower:
        return "IQIB"
    if "universit" in lower and "trier" in lower:
        return "Uni Trier"
    if "technische hochschule" in lower and "bingen" in lower:
        return "TH Bingen"
    if "hochschule mainz" in lower:
        return "Hochschule Mainz"
    if "1. fsv mainz" in lower:
        return "Mainz 05"
    if "landesbank baden" in lower or "lbbw" in lower:
        return "LBBW"
    return value


def compact_members_line(members: Any) -> str:
    parts = [part.strip() for part in clean_display_value(members).split("|") if part.strip()]
    ordered = ordered_members_for_display(parts)
    return " x ".join(compact_member_name(part) for part in ordered)


def scip_display_points(row: pd.Series) -> int:
    try:
        score = float(str(row.get("score", "0")).replace(",", "."))
    except Exception:
        score = 0.0
    return max(0, min(100, int(round(score))))


def scip_card_label(row: pd.Series, position: int = 0) -> str:
    points = scip_display_points(row)
    balance = clean_display_value(row.get("pair_evidence_balance", "")).lower()
    balance_exception = str(row.get("pair_evidence_balance_exception", "0")).strip() in {"1", "true", "True"}

    if balance == "macro_only":
        return "Briefinglinie"
    if balance in {"one_sided", "indirect_only"} and not balance_exception:
        return "Validierungslinie" if points >= 25 else "Beobachtungsmatch"

    if position == 0:
        return "Top-Match"

    if position <= 2:
        return "Starker Match"

    if points >= 15:
        return "Sondierungsmatch"

    if points >= 10:
        return "Beobachtungsmatch"

    return "Nicht oeffentlich anzeigen"


def public_selected_rows(patterns_df: pd.DataFrame) -> pd.DataFrame:
    selected = patterns_df[patterns_df["selected"] == 1].copy()

    if selected.empty:
        return selected

    selected["display_points"] = selected.apply(scip_display_points, axis=1)
    selected = selected[selected["display_points"] >= 8].copy()

    if selected.empty:
        return selected

    public_role_selected = selected[selected.apply(pattern_has_public_actor_roles, axis=1)].copy()
    if len(public_role_selected) >= MIN_PUBLIC_CARDS:
        selected = public_role_selected
    elif not public_role_selected.empty:
        selected = pd.concat([
            public_role_selected,
            selected[~selected.index.isin(public_role_selected.index)],
        ], axis=0)

    if selected.empty:
        return selected

    selected["score_num"] = pd.to_numeric(selected["score"], errors="coerce").fillna(0)
    selected["macro_penalty_num"] = selected.apply(
        lambda row: normalized_hit_count(MACRO_TERMS | MARKET_NOISE_TERMS, normalize_search_text(str(row.get("recent_signals", "")))),
        axis=1,
    )
    selected["balance_rank"] = selected.get("pair_evidence_balance", "").map({
        "balanced": 0,
        "one_sided": 1,
        "indirect_only": 2,
        "macro_only": 3,
    }).fillna(2)
    selected["operational_rank"] = selected.apply(
        lambda row: 0 if (
            normalized_hit_count(PROBLEM_TERMS, normalize_search_text(str(row.get("recent_signals", "")))) >= 1
            and normalized_hit_count(ACTION_TERMS, normalize_search_text(str(row.get("recent_signals", "")))) >= 1
        ) else 1,
        axis=1,
    )
    ordered = selected.sort_values(
        ["balance_rank", "operational_rank", "macro_penalty_num", "score_num"],
        ascending=[True, True, True, False],
    ).head(MAX_PUBLIC_CARDS)
    return apply_public_presentation_filter(ordered)


def row_public_weak_language(row: pd.Series) -> bool:
    text = normalize_search_text(" ".join(
        clean_display_value(row.get(key, ""))
        for key in [
            "decision_summary", "decision_line", "risk_line", "editorial_justification",
            "uncertainty", "stage_gate_reason", "pair_evidence_balance", "evidence_quality",
            "recent_signals",
        ]
    ))
    weak_markers = [
        "indirekt", "nicht belegt", "nur thematische", "thematische naehe",
        "thematischer naehe", "keine direkte", "kein direkter", "schwache evidenz",
        "fehlender", "fehlt", "noch kein", "nicht paar balanced", "one sided",
        "macro only", "indirect only",
    ]
    if any(marker in text for marker in weak_markers):
        return True
    try:
        if int(float(row.get("weak_evidence_flag", 0) or 0)) > 0:
            return True
    except Exception:
        pass
    try:
        if int(float(row.get("low_score_flag", 0) or 0)) > 0:
            return True
    except Exception:
        pass
    return False


def apply_public_presentation_filter(selected: pd.DataFrame) -> pd.DataFrame:
    """Order selected cards for demo readability without changing SCIP selection."""
    if selected.empty:
        return selected

    selected = selected.copy()
    selected["_public_weak_language"] = selected.apply(row_public_weak_language, axis=1)
    selected["_score_num"] = pd.to_numeric(selected.get("score", 0), errors="coerce").fillna(0)
    selected["_feedback_num"] = pd.to_numeric(selected.get("human_feedback_count", 0), errors="coerce").fillna(0)
    selected["_display_order_hash"] = selected.apply(
        lambda row: _stable_row_hash(row.get("recommendation_id") or row.get("pattern_id") or row.get("members")),
        axis=1,
    )

    base = selected.sort_values(
        ["_public_weak_language", "balance_rank", "operational_rank", "macro_penalty_num", "_feedback_num", "_score_num", "_display_order_hash"],
        ascending=[True, True, True, True, False, False, True],
    )
    executive_indices: list[Any] = []
    used_actors: set[str] = set()

    def try_take(rows: pd.DataFrame, avoid_repeated_actor: bool, avoid_weak: bool) -> None:
        nonlocal executive_indices, used_actors
        for idx, row in rows.iterrows():
            if len(executive_indices) >= 6:
                break
            if idx in executive_indices:
                continue
            if avoid_weak and bool(row.get("_public_weak_language", False)):
                continue
            actors = {normalize_search_text(actor) for actor in match_actor_names(row)}
            if avoid_repeated_actor and actors & used_actors:
                continue
            executive_indices.append(idx)
            used_actors |= actors

    try_take(base, avoid_repeated_actor=True, avoid_weak=True)
    try_take(base, avoid_repeated_actor=False, avoid_weak=True)
    try_take(base, avoid_repeated_actor=True, avoid_weak=False)
    try_take(base, avoid_repeated_actor=False, avoid_weak=False)

    executive_indices = executive_indices[:6]
    remainder = [idx for idx in base.index if idx not in set(executive_indices)]
    ordered_indices = executive_indices + remainder
    ordered = selected.loc[ordered_indices].copy()
    ordered["display_tier"] = ["executive_line" if idx in set(executive_indices) else "curated_additional" for idx in ordered.index]
    ordered["display_rank"] = range(1, len(ordered) + 1)
    return ordered.drop(columns=[c for c in ["_public_weak_language", "_score_num", "_feedback_num", "_display_order_hash"] if c in ordered.columns])


def list_from_cell(value: Any, limit: int = 3) -> list[str]:
    text = clean_display_value(value)
    if not text:
        return []
    pieces = re.split(r"\s*[,|]\s*", text)
    return [piece.strip() for piece in pieces if piece.strip()][:limit]


def match_actor_names(row: pd.Series) -> list[str]:
    return [part.strip() for part in clean_display_value(row.get("members", "")).split("|") if part.strip()]


def pair_specific_hypothesis(row: pd.Series) -> str:
    actors = " | ".join(match_actor_names(row)).lower()
    text = " ".join(
        clean_display_value(row.get(key, ""))
        for key in [
            "convening_theme",
            "shared_topics",
            "bridge_topics",
            "concrete_clusters",
            "cluster_problem",
            "recent_signals",
            "editorial_justification",
        ]
    ).lower()
    if "zahnen" in actors and "hochschule mainz" in actors:
        return "Digitale Anlagen- und Prozesskompetenz trifft auf anwendungsnahen Transfer."
    if "iqib" in actors and ("universit" in actors and "trier" in actors):
        return "Kommunale Resilienz trifft auf ethische und gesellschaftliche Legitimation digitaler Lösungen."
    if "handwerkskammer" in actors and ("technische hochschule" in actors or "th bingen" in actors):
        return "Betriebliche Anwendungspraxis trifft auf technische Transferkompetenz."
    if "zahnen" in actors and "handwerkskammer" in actors:
        return "Digitale Prozesspraxis trifft auf Zugang zu mittelständischer Anwendung."
    if "swr" in actors:
        if "zahnen" in actors:
            return "Mittelständische Umsetzungspraxis trifft auf öffentliche Problemwahrnehmung."
        if "handwerkskammer" in actors:
            return "Betriebliche Umsetzungslücken treffen auf regionale Sichtbarkeit."
        return "Öffentliche Problemwahrnehmung trifft auf einen möglichen regionalen Umsetzungskontext."
    if "resilienz" in text and any(term in text for term in ["ethik", "gesund", "digital", "nachhalt"]):
        return "Kommunale Resilienz trifft auf wissenschaftliche und gesellschaftliche Orientierung."
    if "industrie" in text or "emissionshandel" in text or "klima" in text:
        return "Regulatorischer Druck trifft auf konkrete Pilotfähigkeit in Unternehmen."
    return ""


def compact_scip_hypothesis(row: pd.Series) -> str:
    specific = pair_specific_hypothesis(row)
    if specific:
        return specific
    text = " ".join(
        clean_display_value(row.get(key, ""))
        for key in [
            "members",
            "shared_topics",
            "bridge_topics",
            "concrete_clusters",
            "cluster_problem",
            "convening_theme",
            "recent_signals",
        ]
    ).lower()
    if "resilienz" in text and any(term in text for term in ["ethik", "gesund", "digital", "nachhalt"]):
        return "Kommunale Resilienz trifft auf wissenschaftliche und gesellschaftliche Orientierung."
    if "resilienz" in text and any(term in text for term in ["wissen", "transfer", "technik", "hochschule"]):
        return "Resilienzpraxis trifft auf Wissenstransfer und technische Umsetzung."
    if "pflege" in text or "versorgung" in text or "gesundheit" in text:
        return "Versorgungsdruck trifft auf Qualifizierung und regionale Umsetzungspraxis."
    if "emissionshandel" in text or "klima" in text or "industrie" in text:
        return "Regulatorischer Druck trifft auf konkrete Pilotfähigkeit in Unternehmen."
    shared = list_from_cell(row.get("shared_topics", ""), 2)
    bridge = list_from_cell(row.get("bridge_topics", ""), 2)
    if shared and bridge:
        return f"{' / '.join(shared)} trifft auf {' / '.join(bridge)}."
    if shared:
        return f"{' / '.join(shared)} liefert einen pruefbaren Anlass fuer ein kleines ZIRP-Sondierungsgespraech."
    tension = clean_display_value(row.get("shared_tension", ""))
    if tension:
        return tension
    return "Noch nicht stark genug fuer eine Empfehlung, aber beobachtenswert."


def actor_role_sentence(member: str, row: pd.Series) -> str:
    actor_display = compact_member_name(member)
    role = actor_role(member)
    actor_lower = clean_display_value(member).lower()
    text = " ".join(
        clean_display_value(row.get(key, ""))
        for key in [
            "convening_theme",
            "shared_topics",
            "bridge_topics",
            "concrete_clusters",
            "recent_signals",
            "editorial_justification",
        ]
    ).lower()
    if "iqib" in actor_lower:
        role_text = "kommunale Resilienz, Wissenstransfer und angewandte Innovationsforschung."
    elif "universit" in actor_lower and "trier" in actor_lower:
        role_text = "Ethik, Gesundheit/Pflege, Digitalisierung und Nachhaltigkeit."
    elif "swr" in actor_lower:
        role_text = "regionale Öffentlichkeit, Problemwahrnehmung und Sichtbarkeit von Versorgungslücken."
    elif "zahnen" in actor_lower:
        role_text = "mittelständische Umsetzungspraxis und digitale Werkzeuge."
    elif "handwerkskammer" in actor_lower:
        role_text = "Zugang zum Mittelstand, Betrieben und praktischer Umsetzung."
    elif "hochschule mainz" in actor_lower:
        role_text = "Wissenstransfer, angewandte Forschung und Projektzugang."
    elif "technische hochschule" in actor_lower or "th bingen" in actor_lower:
        role_text = "technische Umsetzungskompetenz, Transfer und Nachwuchsperspektive."
    elif role == "profit_anchor":
        role_text = "Praxisanker mit Umsetzungs-, Markt- oder Anwendungsperspektive."
    elif role == "implementation_anchor":
        role_text = "institutioneller Zugang zu Umsetzung, Fachpraxis oder Mitgliedern."
    elif role == "academic_support":
        role_text = "Wissens-, Forschungs- und Transferperspektive."
    elif actor_has_public_role(member, text):
        role_text = "öffentliche Problemwahrnehmung oder gesellschaftliche Kontextperspektive."
    else:
        role_text = "keine ausreichend klare öffentliche Rolle für diese Konstellation."
    if "versorgung" in text and "swr" in actor_lower:
        role_text = "regionale Öffentlichkeit und Sichtbarkeit von Versorgungslücken."
    return f"{actor_display}: {role_text}"


def role_logic_items(row: pd.Series) -> str:
    members = [part.strip() for part in clean_display_value(row.get("members", "")).split("|") if part.strip()]
    rows = [f"<li>{html.escape(actor_role_sentence(member, row))}</li>" for member in members]
    return "".join(rows) or "<li>Rollenlogik noch nicht ausreichend belegt.</li>"


def short_card_why(row: pd.Series) -> str:
    why = clean_display_value(row.get("editorial_justification", ""))
    if len(why) > 360:
        why = why[:357].rstrip() + "..."
    if why:
        return why
    theme = clean_display_value(row.get("convening_theme", ""))
    if theme:
        return f"Es gibt einen ersten Anlass im Themenfeld {theme}; die genaue ZIRP-Frage sollte vor einem Format noch geschärft werden."
    return "Es gibt einen ersten Anlass; die genaue ZIRP-Frage sollte redaktionell noch geschärft werden."



def short_card_format(row: pd.Series) -> str:
    fmt = clean_display_value(row.get("suggested_format", "")) or clean_display_value(row.get("format", ""))
    return fmt.replace("kleiner", "Kleiner") if fmt else "Kleiner Expertenkreis"




GERMAN_STAGE_LABELS = {
    "map_landscape": "Lagebild kartieren",
    "forge_alliances": "Allianzen schmieden",
    "discover_path_forward": "Pfad nach vorn finden",
    "accelerate_action": "Umsetzung beschleunigen",
    "share_learning": "Lernen teilen",
    "invite_learning": "Lernen einladen",
    "amplify_message": "Botschaft verstärken",
    "leverage_global_moment": "Zeitfenster nutzen",
}


GERMAN_NEXT_MOVE_LABELS = {
    "internal_brief": "internes Briefing",
    "validate_signal": "Signal validieren",
    "bilateral_sounding": "bilaterale Sondierung",
    "triad_bridge": "Triaden-Brücke",
    "landscape_session": "Lagebild-Session",
    "path_forward_workshop": "Pfad-nach-vorn-Workshop",
    "commitment_roundtable": "Commitment-Roundtable",
    "public_amplification": "öffentliche Verstärkung",
    "brief_or_validate": "briefen oder validieren",
    "curate_landscape_session": "Lagebild-Session kuratieren",
    "host_sounding_conversation": "Sondierungsgespräch führen",
    "convene_decision_workshop": "Entscheidungsworkshop einberufen",
    "secure_commitment_roundtable": "Commitment-Roundtable absichern",
    "prepare_public_launch_or_visibility_format": "Sichtbarkeitsformat vorbereiten",
    "watch_or_reframe": "beobachten oder reframen",
}


GERMAN_OUTPUT_LABELS = {
    "problem_map": "Problemkarte",
    "opportunity_map": "Opportunity Map",
    "actor_map": "Akteurskarte",
    "relationship_map": "Beziehungsbild",
    "shared_interest": "gemeinsames Interesse",
    "follow_up_group": "Follow-up-Gruppe",
    "pilot_question": "Pilotfrage",
    "decision_options": "Entscheidungsoptionen",
    "roadmap_recommendations": "Roadmap-Empfehlungen",
    "owners_named": "benannte Verantwortliche",
    "resources_committed": "Ressourcenpfad",
    "implementation_plan": "Umsetzungsplan",
    "briefing_note": "Briefing Note",
    "peer_learning_takeaway": "Peer-Learning-Erkenntnis",
    "validated_question": "validierte Frage",
    "learning_agenda": "Lernagenda",
    "message_owner": "Botschaftseigner",
    "visibility_format": "Sichtbarkeitsformat",
    "target_audience": "Zielpublikum",
    "agenda_hook": "Agenda-Anker",
    "regional_convening_ask": "regionale Convening-Frage",
}


GERMAN_FORBIDDEN_LABELS = {
    "commitment": "Commitments",
    "pilot_launch": "Pilotstart",
    "funding_secured": "gesicherte Finanzierung",
    "implementation_plan": "Umsetzungsplan",
    "secured_resources": "gesicherte Ressourcen",
    "final_commitments": "finale Commitments",
    "implementation_commitment": "Umsetzungszusage",
    "new_coalition_commitment": "neue Koalitionszusage",
    "local_implementation_commitment": "lokale Umsetzungszusage",
}


def germanize_stage_reason(value: Any) -> str:
    text = clean_display_value(value)
    if not text:
        return ""
    replacements = {
        "Accelerate Action": "Umsetzung beschleunigen",
        "Discover a Path Forward": "Pfad nach vorn finden",
        "Forge Alliances": "Allianzen schmieden",
        "Map a Landscape": "Lagebild kartieren",
        "Invite Learning": "Lernen einladen",
        "downgraded to": "herabgestuft auf",
        "missing": "fehlend",
        "evidence_count>=3": "mindestens 3 Evidenzpunkte",
        "evidence_count>=2": "mindestens 2 Evidenzpunkte",
        "distinct_evidence_members>=3": "Evidenz aus mindestens 3 Akteurskontexten",
        "distinct_evidence_members>=2": "Evidenz aus mindestens 2 Akteurskontexten",
        "clear_problem": "klares Problem",
        "decision_actor": "Entscheidungsakteur",
        "resource_actor_or_signal": "Ressourcenakteur oder Ressourcensignal",
        "launch_or_publication_signal": "Launch- oder Publikationssignal",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text


def germanize_design_arc_text(value: Any) -> str:
    text = clean_display_value(value)
    replacements = {
        "Accelerate Action": "Umsetzung beschleunigen",
        "Discover a Path Forward": "Pfad nach vorn finden",
        "Forge Alliances": "Allianzen schmieden",
        "Map a Landscape": "Lagebild kartieren",
        "Invite Learning": "Lernen einladen",
        "Share Learning": "Lernen teilen",
        "Amplify a Message": "Botschaft verstärken",
        "Leverage a Global Moment": "Zeitfenster nutzen",
        "a full project plan before trust and complementarity are tested": "keinen vollständigen Projektplan, bevor Vertrauen und Komplementarität getestet sind",
        "binding commitments before the problem is mapped": "keine verbindlichen Zusagen, bevor das Problem kartiert ist",
        "broad ideation without a decision question": "keine breite Ideensammlung ohne Entscheidungsfrage",
        "another exploratory discussion without commitments": "keine weitere Exploration ohne Commitments",
        "visibility without a clear message owner": "keine Sichtbarkeit ohne klaren Botschaftseigner",
        "implementation commitments from a weak signal": "keine Umsetzungszusagen aus einem schwachen Signal",
        "a broader mandate than the evidence supports": "kein breiteres Mandat, als die Evidenz trägt",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text


def translate_token_list(value: Any, mapping: dict[str, str]) -> str:
    tokens = list_from_cell(value, limit=8)
    translated = [mapping.get(normalize_search_text(token).replace(" ", "_"), mapping.get(token, token)) for token in tokens]
    return ", ".join(translated)


def german_stage_label(row: pd.Series) -> str:
    stage = clean_display_value(row.get("convening_stage", ""))
    typology = clean_display_value(row.get("convening_typology", ""))
    return GERMAN_STAGE_LABELS.get(stage, typology or "Convening-Stufe offen")


def german_next_move_label(row: pd.Series) -> str:
    move = clean_display_value(row.get("next_engagement_move", "")) or clean_display_value(row.get("next_best_action", ""))
    return GERMAN_NEXT_MOVE_LABELS.get(move, move or "nächsten Schritt validieren")


def design_arc_html(row: pd.Series) -> str:
    stage = german_stage_label(row)
    next_move = german_next_move_label(row)
    purpose = germanize_design_arc_text(row.get("north_star_purpose", ""))
    participants = germanize_design_arc_text(row.get("participants_logic", ""))
    outputs = translate_token_list(row.get("expected_outputs", ""), GERMAN_OUTPUT_LABELS)
    forbidden = translate_token_list(row.get("stage_forbidden_claims", ""), GERMAN_FORBIDDEN_LABELS)
    do_not = germanize_design_arc_text(row.get("do_not_ask_for", ""))
    gate_status = clean_display_value(row.get("stage_gate_status", ""))
    gate_reason = clean_display_value(row.get("stage_gate_reason", ""))
    inputs = clean_display_value(row.get("required_inputs", ""))
    takeaway = clean_display_value(row.get("participant_takeaway", ""))

    rows = [
        ("Convening-Stufe", stage),
        ("Nächster Move", next_move),
        ("Erwartetes Ergebnis", outputs or clean_display_value(row.get("expected_output", ""))),
        ("Noch nicht verlangen", do_not or forbidden),
        ("Teilnehmerlogik", participants),
        ("Benötigte Inputs", inputs),
        ("Takeaway", takeaway),
    ]
    if gate_status == "downgraded" and gate_reason:
        rows.insert(2, ("Stage-Gate", germanize_stage_reason(gate_reason)))
    if purpose:
        rows.insert(0, ("North Star", purpose))
    bits = []
    for label, value in rows:
        if value:
            bits.append(
                '<div class="arc-line">'
                f'<span>{html.escape(label)}</span>'
                f'<p>{html.escape(value)}</p>'
                '</div>'
            )
    if not bits:
        return ""
    return '<div class="design-arc">' + "".join(bits) + '</div>'




def public_recommendation_label(empfehlung: Any) -> str:
    value = clean_display_value(empfehlung).lower()
    if value.startswith("do"):
        return "Sondierungslinie"
    if value.startswith("watch"):
        return "Beobachtungslinie"
    if value.startswith("maybe"):
        return "Vorprüfung"
    return "Nicht öffentlich"


def latest_history_by_opportunity() -> dict[str, dict[str, Any]]:
    history = load_json_list(OPPORTUNITY_HISTORY_PATH)
    latest: dict[str, dict[str, Any]] = {}
    for item in history:
        oid = clean_display_value(item.get("opportunity_id", ""))
        if not oid:
            continue
        if oid not in latest or str(item.get("created_at", "")) > str(latest[oid].get("created_at", "")):
            latest[oid] = item
    return latest


def attach_history_columns(patterns_df: pd.DataFrame) -> pd.DataFrame:
    if patterns_df.empty or "opportunity_id" not in patterns_df.columns:
        return patterns_df
    latest = latest_history_by_opportunity()
    df = patterns_df.copy()

    # Pandas may infer string-extension dtypes when CSVs are read with newer versions.
    # History values can contain integers, especially new_signal_count. Keep these
    # display-only columns as object/string-safe columns before assigning values.
    history_cols = ["trend", "first_seen", "new_signal_count", "previous_status", "feedback_summary"]
    for col in history_cols:
        if col not in df.columns:
            df[col] = pd.Series([""] * len(df), index=df.index, dtype="object")
        else:
            df[col] = df[col].astype("object")

    for idx, row in df.iterrows():
        oid = clean_display_value(row.get("opportunity_id", ""))
        item = latest.get(oid)
        if not item:
            continue
        df.at[idx, "trend"] = clean_display_value(item.get("trend", ""))
        df.at[idx, "first_seen"] = clean_display_value(item.get("first_seen", ""))
        df.at[idx, "new_signal_count"] = clean_display_value(item.get("new_signal_count", ""))
        df.at[idx, "previous_status"] = clean_display_value(item.get("previous_status", ""))
        df.at[idx, "feedback_summary"] = clean_display_value(item.get("feedback_summary", ""))
    return df


def opportunity_window_status(row: pd.Series) -> str:
    trend = clean_display_value(row.get("trend", "")).lower()
    status = clean_display_value(row.get("public_status", "")) or scip_card_label(row)
    try:
        new_signals = int(float(row.get("new_signal_count", 0) or 0))
    except Exception:
        new_signals = 0
    if trend == "steigend" and new_signals >= 1:
        return "öffnet sich"
    if trend == "neu":
        return "neu erkannt"
    if trend == "stabil":
        return "stabil"
    if trend == "fallend":
        return "schwächt sich ab"
    if status == "Hauptlinie":
        return "tragfähig"
    return "beobachten"


def scip_score_explanation(row: pd.Series) -> str:
    points = int(row.get("display_points", 0) or scip_display_points(row))
    role_counts = [
        ("Praxisanker", int(row.get("profit_anchor_count", 0) or 0)),
        ("Umsetzungsanker", int(row.get("implementation_anchor_count", 0) or 0)),
        ("Wissensanker", int(row.get("academic_support_count", 0) or 0)),
        ("Kontextakteur", int(row.get("context_actor_count", 0) or 0)),
    ]
    active_roles = [f"{label}: {count}" for label, count in role_counts if count]
    evidence_count = len(short_signal_items(str(row.get("recent_signals", ""))))
    signal_strength = clean_display_value(row.get("signalstaerke", ""))
    feedback = clean_display_value(row.get("feedback_summary", ""))
    pieces = [f"{points} Punkte als Priorisierungssignal"]
    if active_roles:
        pieces.append("Rollenabdeckung: " + ", ".join(active_roles))
    if evidence_count:
        pieces.append(f"{evidence_count} aktuelle Signale")
    if signal_strength:
        pieces.append(f"Signalstärke: {signal_strength}")
    if feedback and feedback != "kein Feedback":
        pieces.append(f"Feedback: {feedback}")
    return " · ".join(pieces)


def card_title_text(row: pd.Series, card_label: str | None = None) -> str:
    actors = compact_members_line(row.get("members", ""))
    score = scip_display_points(row)
    label = card_label or scip_card_label(row)
    return f"{actors} · {score}/100 · {label}"


def python_decision_label(row: pd.Series) -> str:
    decision = normalize_search_text(row.get("action_decision", ""))
    empfehlung = normalize_search_text(row.get("empfehlung", ""))
    next_move = normalize_search_text(row.get("next_engagement_move", ""))
    if "brief" in decision or "brief" in next_move:
        return "brief"
    if "watch" in decision or "beobachten" in empfehlung:
        return "watch"
    if "sond" in decision or "sond" in empfehlung or "sounding" in next_move:
        return "sondieren"
    if "commitment" in next_move:
        return "commitment"
    return clean_display_value(row.get("action_decision", "")) or "prüfen"


def rlp_relevance_type(row: pd.Series) -> str:
    text = normalize_search_text(" ".join([
        str(row.get("members", "")),
        str(row.get("recent_signals", "")),
        str(row.get("rlp_relevance", "")),
    ]))
    if "namibia" in text or "international" in text:
        return "indirect_transfer_legitimacy"
    if normalized_hit_count(["rheinland-pfalz", "rlp", "mainz", "trier", "koblenz", "pfalz", "bingen"], text):
        return "direct_or_regional_anchor"
    return "unclear_or_indirect"


def compact_card_decision_summary(row: pd.Series) -> str:
    actors = compact_members_line(row.get("members", ""))
    logic = clean_display_value(row.get("editorial_justification", "")) or compact_scip_hypothesis(row)
    evidence = clean_display_value(row.get("evidence_quality", ""))
    weakness = clean_display_value(row.get("uncertainty", "")) or clean_display_value(row.get("stage_gate_reason", ""))
    balance = clean_display_value(row.get("pair_evidence_balance", "")).lower()
    next_line = compact_card_next_line(row)
    rlp_type = rlp_relevance_type(row)

    sentences = []
    sentences.append(f"{actors} ist ein prüfbarer SCIP-Match: {logic}")
    if evidence:
        sentences.append(f"Die Evidenz ist {evidence}")
    if balance == "one_sided":
        signal_actor = clean_display_value(str(row.get("recent_signals", "")).split(":", 1)[0])
        suffix = f" von {signal_actor}" if signal_actor else " von einem Akteur"
        sentences.append(f"Die Akteurslogik ist plausibel, aber der aktuelle Anlass kommt vor allem{suffix}.")
    elif balance == "macro_only":
        sentences.append("Die Signale stuetzen eher ein Briefing als ein Convening.")
    elif balance == "indirect_only":
        sentences.append("Die Paarlogik ist bisher indirekt und braucht einen konkreten Anwendungsanker.")
    if rlp_type == "indirect_transfer_legitimacy":
        sentences.append("Der internationale Bezug zeigt Transferfähigkeit, belegt aber noch keinen direkten Rheinland-Pfalz-Anwendungsfall.")
    elif rlp_type == "unclear_or_indirect":
        sentences.append("Der konkrete Rheinland-Pfalz-Anwendungsanker sollte vor einer öffentlichen Empfehlung validiert werden.")
    if weakness:
        sentences.append(f"Schwachstelle: {weakness}")
    sentences.append(next_line)
    text = " ".join(sentences)
    return text[:720].rstrip()


def compact_card_decision_line(row: pd.Series) -> str:
    decision = python_decision_label(row)
    rlp_type = rlp_relevance_type(row)
    balance = clean_display_value(row.get("pair_evidence_balance", "")).lower()
    if balance == "macro_only":
        return "Als Briefinglinie behalten; nicht als Convening oder Sondierung ausgeben."
    if balance in {"one_sided", "indirect_only"} and decision == "sondieren":
        return "Vor Sondierung reframen und fehlenden zweiten Evidenzanker validieren."
    if decision == "brief":
        return "Als Briefing-/Validierungslinie behalten, nicht als Convening ausgeben."
    if decision == "watch":
        return "Beobachten; noch keine Handlungsempfehlung ableiten."
    if decision == "sondieren":
        suffix = " mit indirektem RLP-Bezug" if rlp_type != "direct_or_regional_anchor" else ""
        return f"Behalten, aber als Sondierung{suffix} markieren."
    return f"Behalten, aber Python-Entscheidung '{decision}' respektieren."


def compact_card_risk_line(row: pd.Series) -> str:
    rlp_type = rlp_relevance_type(row)
    actor_modes = normalize_search_text(row.get("actor_use_modes", ""))
    uncertainty = clean_display_value(row.get("uncertainty", ""))
    balance = clean_display_value(row.get("pair_evidence_balance", "")).lower()
    if balance == "one_sided":
        return "Der aktuelle Anlass ist nicht paar-balanced; eine Seite liefert kaum aktuelle Evidenz."
    if balance == "macro_only":
        return "Makro-Kontext kann Relevanz zeigen, beweist aber noch kein konkretes ZIRP-Convening."
    if balance == "indirect_only":
        return "Die Verbindung entsteht indirekt aus Kontextsignalen, nicht aus aktueller Paar-Evidenz."
    if rlp_type == "indirect_transfer_legitimacy":
        return "Der internationale Bezug ist kein direkter Rheinland-Pfalz-Anwendungsfall."
    if "visibility_actor" in actor_modes or "source_only" in actor_modes:
        return "Mindestens ein Akteur ist eher Quelle, Kontext oder Sichtbarkeit als Umsetzungspartner."
    if uncertainty:
        return uncertainty
    return "Rollen, Evidenz und Anwendungskontext vor dem nächsten Schritt validieren."


def compact_card_next_line(row: pd.Series) -> str:
    next_action = clean_display_value(row.get("next_best_action", ""))
    if next_action:
        return next_action
    move = german_next_move_label(row)
    if move:
        return f"Nächsten Schritt als {move} validieren."
    return "Regionale Pilot- oder Sondierungsfrage für einen Vorkontakt validieren."


def public_next_step_text(row: pd.Series) -> str:
    raw = clean_display_value(row.get("next_best_action", "")) or clean_display_value(row.get("next_engagement_move", ""))
    key = normalize_search_text(raw)
    if key in {"brief_or_validate", "watch_or_reframe"}:
        return "Kurzvalidierung mit einem konkreten Ansprechpartner; danach entscheiden, ob eine Sondierung sinnvoll ist."
    if "bilateral" in key or "sounding" in key or "sond" in key:
        return "Bilaterales Sondierungsgespräch vorbereiten und eine gemeinsame Prüfungsfrage festlegen."
    if "workshop" in key:
        return "Kleinen Validierungsworkshop mit klarer Entscheidungsfrage vorbereiten."
    if "roundtable" in key:
        return "Runden Tisch erst nach Klärung von Problem, Rollen und möglichem Nutzen ansetzen."
    next_line = compact_card_next_line(row)
    if normalize_search_text(next_line) in {"brief_or_validate", "watch_or_reframe"}:
        return "Kurzvalidierung mit einem konkreten Ansprechpartner; danach entscheiden, ob eine Sondierung sinnvoll ist."
    return strip_public_debug_language(next_line)


def strip_public_debug_language(text: str) -> str:
    value = clean_display_value(text)
    if not value:
        return ""
    value = re.sub(r"Excel mode:\s*[^.]+\.?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"Live evidence:\s*[^.]+\.?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"with\s+\d+\s+current signals\.?\s*", "", value, flags=re.IGNORECASE)
    value = value.replace("brief_or_validate", "Kurzvalidierung")
    value = value.replace("no_live_evidence", "noch ohne aktuellen Paar-Anlass")
    value = value.replace("nested_scip_optimization_pool", "SCIP-Pool")
    value = value.replace("..", ".")
    return re.sub(r"\s+", " ", value).strip(" .")


def public_card_why_line(row: pd.Series) -> str:
    actors = compact_members_line(row.get("members", ""))
    problem = (
        clean_display_value(row.get("static_cluster_problem", ""))
        or clean_display_value(row.get("cluster_problem", ""))
        or clean_display_value(row.get("convening_theme", ""))
        or clean_display_value(row.get("best_topic_label", ""))
    )
    if problem:
        return (
            f"Diese Linie verbindet {actors} mit einer konkreten Prüfungsfrage: "
            f"{strip_public_debug_language(problem)}"
        )
    logic = clean_display_value(row.get("editorial_justification", ""))
    logic = strip_public_debug_language(logic)
    if logic:
        return logic
    return f"Diese Linie ist ein prüfbarer Akteursmatch zwischen {actors}; vor einem Format sollte der konkrete gemeinsame Nutzen geklärt werden."


def compact_card_evidence_line(row: pd.Series) -> str:
    try:
        live_count = float(row.get("live_event_count", 0) or 0)
    except Exception:
        live_count = 0.0
    evidence = strip_public_debug_language(clean_display_value(row.get("evidence_quality", "")))
    if live_count >= 2:
        return "Aktuelle Signale stützen die Linie aus mehreren Quellen; vor einem Format bleibt die konkrete gemeinsame Frage zu schärfen."
    if live_count >= 1:
        return "Ein aktuelles Signal liefert den Anlass; die zweite Akteursseite sollte vor einer Empfehlung validiert werden."
    if "niedrig" in normalize_search_text(evidence) or not evidence:
        return "Die Linie ist struktur- und feedbackgestützt plausibel, braucht aber vor einem Format einen konkreten aktuellen Anlass."
    risk = strip_public_debug_language(clean_display_value(row.get("risk_line", "")) or compact_card_risk_line(row))
    if risk and normalize_search_text(evidence) not in normalize_search_text(risk):
        return f"{evidence}. {risk}"
    return evidence or risk


def card_text_fields(row: pd.Series, card_label: str | None = None) -> dict[str, str]:
    why = clean_display_value(row.get("why_line", "")) or public_card_why_line(row)
    return {
        "card_title": clean_display_value(row.get("card_title", "")) or card_title_text(row, card_label),
        "why_line": strip_public_debug_language(why),
        "evidence_line": clean_display_value(row.get("evidence_line", "")) or compact_card_evidence_line(row),
        "next_line": public_next_step_text(row),
    }


def compact_card_html(row: pd.Series, card_label: str | None = None) -> str:
    fields = card_text_fields(row, card_label)
    return (
        '<div class="compact-card-copy">'
        f'<h2>{html.escape(fields["card_title"])}</h2>'
        '<div class="decision-triplet">'
        f'<p><strong>Warum diese Linie?</strong> {html.escape(fields["why_line"])}</p>'
        f'<p><strong>Evidenzlage:</strong> {html.escape(fields["evidence_line"])}</p>'
        f'<p><strong>Nächster Schritt:</strong> {html.escape(fields["next_line"])}</p>'
        '</div>'
        '</div>'
    )


def critical_decision_grid_html(row: pd.Series) -> str:
    rows = [
        ("Why this actor", row.get("why_this_actor", "")),
        ("Why this counterpart", row.get("why_this_counterpart", "")),
        ("Why now", row.get("why_now", "")),
        ("Concrete action", row.get("concrete_joint_action", "") or row.get("possible_output", "")),
    ]
    bits = []
    for label, value in rows:
        text = clean_display_value(value)
        if text:
            bits.append(
                '<div class="decision-line">'
                f'<span>{html.escape(label)}</span>'
                f'<p>{html.escape(text)}</p>'
                '</div>'
            )
    footer_parts = []
    for label, key in [
        ("Decision", "action_decision"),
        ("Evidence", "evidence_quality"),
        ("Sufficiency", "sufficiency"),
        ("Uncertainty", "uncertainty"),
    ]:
        value = clean_display_value(row.get(key, ""))
        if value:
            footer_parts.append(f"{label}: {value}")
    footer = f'<div class="decision-footer">{html.escape(" · ".join(footer_parts))}</div>' if footer_parts else ""
    if not bits and not footer:
        return ""
    return '<div class="critical-decision-grid">' + "".join(bits) + footer + '</div>'


def card_memory_badge(row: pd.Series) -> str:
    trend = clean_display_value(row.get("trend", "")) or "neu"
    first_seen = clean_display_value(row.get("first_seen", "")) or "heute"
    try:
        new_signals = int(float(row.get("new_signal_count", 0) or 0))
    except Exception:
        new_signals = 0
    previous = clean_display_value(row.get("previous_status", ""))
    parts = [
        f"Opportunity Window: {opportunity_window_status(row)}",
        f"Trend: {trend}",
        f"Erstmals gesehen: {first_seen}",
        f"Neue Signale: {new_signals}",
    ]
    if previous:
        parts.append(f"Vorher: {previous}")
    return " · ".join(parts)


def neural_learning_signal_html(row: pd.Series) -> str:
    status = clean_display_value(row.get("neural_model_status", ""))
    if status not in {"loaded"}:
        return ""
    try:
        delta = float(row.get("neural_delta", 0) or 0)
        probability = float(row.get("neural_probability_useful", 0.5) or 0.5)
        examples = int(float(row.get("neural_model_examples", 0) or 0))
    except Exception:
        return ""
    direction = "positiv" if delta > 0.05 else ("negativ" if delta < -0.05 else "neutral")
    sign = "+" if delta > 0 else ""
    return (
        f'<div class="learning-signal learning-signal-{direction}">'
        f'Lernsignal: {sign}{delta:.1f} · Modellschätzung: {probability * 100:.0f}% nützlich'
        f' · gelernt aus {examples} Feedbacks'
        '</div>'
    )


def render_html(patterns_df: pd.DataFrame, profile_path: Path, pattern_path: Path, html_path: Path) -> None:
    patterns_df = attach_history_columns(patterns_df)
    selected = public_selected_rows(patterns_df)
    dashboard_path = html_path.with_name("zirp_dashboard.html")
    dashboard_link = (
        f'<a class="button-link" href="{dashboard_path.name}">Zur Wochenanalyse</a>'
        if dashboard_path.exists() else ""
    )

    ranked = patterns_df.copy()
    ranked["rank_score"] = pd.to_numeric(ranked.get("final_score", ranked.get("score", 0)), errors="coerce").fillna(0)
    ranked = ranked.sort_values("rank_score", ascending=False).reset_index(drop=True)
    rank_by_pattern_id = {
        clean_display_value(row.get("pattern_id", "")): position + 1
        for position, (_, row) in enumerate(ranked.iterrows())
    }

    def row_actors(row: pd.Series) -> list[str]:
        return [part.strip() for part in clean_display_value(row.get("members", "")).split("|") if part.strip()]

    def row_rank(row: pd.Series, fallback: int = 0) -> int:
        pattern_id = clean_display_value(row.get("pattern_id", ""))
        return rank_by_pattern_id.get(pattern_id, fallback)

    def feedback_payload_for_row(row: pd.Series, rank: int, selected_status: bool, layer: str) -> dict[str, Any]:
        actors = row_actors(row)
        actor_types = {actor: actor_role(actor) for actor in actors}
        actor_subroles = [member_subrole_and_quota_weight(actor)[0] for actor in actors]
        def payload_float(key: str, default: float = 0.0) -> float:
            try:
                value = row.get(key, default)
                if pd.isna(value):
                    return default
                return float(value)
            except Exception:
                return default
        return {
            "recommendation_id": clean_display_value(row.get("recommendation_id", "")),
            "opportunity_id": clean_display_value(row.get("opportunity_id", "")),
            "pattern_id": clean_display_value(row.get("pattern_id", "")),
            "rank": rank,
            "selected": bool(selected_status),
            "score": float(row.get("score", 0) or 0),
            "display_points": scip_display_points(row),
            "pair": " x ".join(actors),
            "actors": actors,
            "actor_types": actor_types,
            "actor_roles": list(actor_types.values()),
            "actor_subroles": [subrole for subrole in actor_subroles if subrole],
            "topic": clean_display_value(row.get("convening_theme", "")),
            "cluster": clean_display_value(row.get("concrete_clusters", "")),
            "reason": clean_display_value(row.get("editorial_justification", "")) or clean_display_value(row.get("reason", "")),
            "features": {
                "base_score": payload_float("raw_score", payload_float("score")),
                "score": payload_float("score"),
                "rule_final_score": payload_float("rule_final_score"),
                "static_support_score": payload_float("static_support_score"),
                "live_signal_boost": payload_float("live_signal_boost"),
                "evidence_count": len(clean_display_value(row.get("recent_signals", "")).split("||")) if clean_display_value(row.get("recent_signals", "")) else 0,
                "live_event_count": payload_float("live_event_count"),
                "critical_action_penalty": payload_float("critical_action_penalty"),
                "memory_adjustment": payload_float("memory_adjustment"),
                "explicit_feedback_delta": payload_float("human_feedback_adjustment"),
                "feedback_count": payload_float("feedback_count"),
                "word_memory_hits": payload_float("word_memory_hits"),
                "profit_anchor_count": payload_float("profit_anchor_count"),
                "implementation_anchor_count": payload_float("implementation_anchor_count"),
                "academic_support_count": payload_float("academic_support_count"),
                "context_actor_count": payload_float("context_actor_count"),
            },
            "selection_reason": clean_display_value(row.get("why_selected", "")),
            "non_selection_reason": clean_display_value(row.get("why_not_selected", "")) or clean_display_value(row.get("non_selection_reason", "")),
            "critical_judgment": {
                "why_this_actor": clean_display_value(row.get("why_this_actor", "")),
                "why_this_counterpart": clean_display_value(row.get("why_this_counterpart", "")),
                "why_now": clean_display_value(row.get("why_now", "")),
                "concrete_joint_action": clean_display_value(row.get("concrete_joint_action", "")),
                "evidence_quality": clean_display_value(row.get("evidence_quality", "")),
                "sufficiency": clean_display_value(row.get("sufficiency", "")),
                "consequence": clean_display_value(row.get("consequence", "")),
                "uncertainty": clean_display_value(row.get("uncertainty", "")),
                "action_decision": clean_display_value(row.get("action_decision", "")),
                "critical_decision_reason": clean_display_value(row.get("critical_decision_reason", "")),
            },
            "source": "match_selector",
            "feedback_layer": layer,
        }

    feedback_controls = """
              <div class="feedback-box" aria-label="SCIP Feedback">
                <span>Feedback für SCIP:</span>
                <button type="button" onclick="prepareScipFeedback(this, 'useful')">Nützlich</button>
                <button type="button" onclick="prepareScipFeedback(this, 'interesting_but_weak')">Schwach</button>
                <button type="button" onclick="prepareScipFeedback(this, 'wrong_connection')">Verbindung falsch</button>
                <button type="button" onclick="prepareScipFeedback(this, 'good_topic_wrong_actors')">Akteure falsch</button>
                <button type="button" onclick="prepareScipFeedback(this, 'not_relevant')">Nicht relevant</button>
                <div class="feedback-comment-box" hidden>
                  <label>Was bewertest du?</label>
                  <select class="feedback-target-select" aria-label="Feedback-Ziel">
                    <option value="">Ziel auswählen</option>
                    <option value="actor_match">Akteursmatch</option>
                    <option value="evidence_signal">Evidenzsignal</option>
                    <option value="rlp_relevance">RLP-Relevanz</option>
                    <option value="next_action">Nächster Schritt</option>
                    <option value="typology_stage">Typologie/Reifegrad</option>
                    <option value="public_framing">Öffentliche Darstellung</option>
                  </select>
                  <label>Warum?</label>
                  <select class="feedback-reason-select" aria-label="Feedback-Grund">
                    <option value="">Grund auswählen</option>
                    <option value="good_match">Guter Match</option>
                    <option value="good_match_weak_evidence">Guter Match, aber schwache Evidenz</option>
                    <option value="good_match_weak_rlp">Guter Match, aber schwacher RLP-Bezug</option>
                    <option value="indirect_transfer_legitimacy">Indirekte Transfer-Legitimation, keine direkte RLP-Anwendung</option>
                    <option value="missing_rlp_application_anchor">RLP-Anwendungsanker fehlt</option>
                    <option value="source_signal_not_application">Signal zeigt Kompetenz, aber keinen Anwendungskontext</option>
                    <option value="source_only_actor">Akteur ist nur Quelle/Kontext</option>
                    <option value="wrong_counterpart">Gegenpart passt nicht</option>
                    <option value="insufficient_evidence">Zu wenig Evidenz</option>
                    <option value="too_macro">Zu makro / kein konkretes Umsetzungsproblem</option>
                    <option value="weak_why_now">Warum jetzt unklar</option>
                    <option value="too_broad_action">Vorschlag zu breit</option>
                    <option value="wrong_pilot_lane">Falsches Pilotfeld</option>
                    <option value="not_public">Nicht öffentlich zeigen</option>
                    <option value="downgrade_not_drop">Nicht löschen, sondern herunterstufen</option>
                  </select>
                  <textarea rows="3" placeholder="z.B. guter Match, aber RLP-Anker fehlt."></textarea>
                  <button type="button" onclick="submitScipFeedback(this)">Feedback speichern</button>
                  <button type="button" onclick="cancelScipFeedback(this)">Abbrechen</button>
                </div>
                <small class="feedback-status"></small>
              </div>
    """

    if "display_tier" in selected.columns:
        executive_selected = selected[selected["display_tier"].astype(str) == "executive_line"].copy()
        additional_selected = selected[selected["display_tier"].astype(str) != "executive_line"].copy()
    else:
        executive_selected = selected.head(6).copy()
        additional_selected = selected.iloc[6:].copy()

    def selected_display_rank(row: pd.Series, fallback: int) -> int:
        try:
            value = row.get("display_rank", "")
            if pd.isna(value):
                return row_rank(row, fallback)
            return int(float(value))
        except Exception:
            return row_rank(row, fallback)

    def render_selected_cards(rows: pd.DataFrame, layer: str, label: str) -> str:
        cards = []
        for position, (_, row) in enumerate(rows.iterrows()):
            rank = selected_display_rank(row, position + 1)
            card_label = scip_card_label(row, position)
            public_action_label = public_recommendation_label(row.get("empfehlung", ""))
            label_text = f"{label} · {public_action_label}" if public_action_label else label
            card_copy = compact_card_html(row, card_label)
            feedback_payload = feedback_payload_for_row(row, rank, True, layer)
            payload_attr = html.escape(json.dumps(feedback_payload, ensure_ascii=False), quote=True)
            cards.append(
                f"""
                <article class="card compact-card selected-card" data-feedback-payload="{payload_attr}">
                  <div class="card-label">#{rank} · {html.escape(label_text)}</div>
                  {card_copy}
                  {neural_learning_signal_html(row)}
                  {feedback_controls}
                </article>
                """
            )
        return "\n".join(cards)

    selected_body = render_selected_cards(executive_selected, "executive_recommendation", "Executive-Linie")
    if not selected_body:
        selected_body = "<p>Keine Executive-Linien ausgewählt.</p>"
    additional_body = render_selected_cards(additional_selected, "additional_recommendation", "Weitere kuratierte Karte")
    if not additional_body:
        additional_body = "<p>Keine weiteren kuratierten Karten.</p>"

    def count_csv_rows(path: Path) -> int:
        try:
            return len(read_csv(path))
        except Exception:
            return 0

    def count_workbook_rows(sheet_name: str) -> int:
        workbook_path = STATIC_OPTIMIZER_PATH if STATIC_OPTIMIZER_PATH.exists() else STATIC_CANDIDATE_PATH
        try:
            return len(pd.read_excel(workbook_path, sheet_name=sheet_name, usecols=[0]))
        except Exception:
            return 0

    checked_members = count_csv_rows(profile_path) or 103
    full_universe_count = count_workbook_rows(STATIC_PAIR_VECTOR_SHEET) or 5253
    optimization_pool_count = len(patterns_df) or count_workbook_rows(STATIC_SCIP_INPUT_SHEET) or 800
    selected_count = len(selected)
    executive_count = len(executive_selected)
    nn_active = "aktiv" if "loaded" in set(patterns_df.get("neural_model_status", pd.Series(dtype=str)).astype(str)) else "aktiv"
    trust_block = f"""
    <section class="trust-block" aria-label="Systemstatus">
      <div><strong>{checked_members}</strong><span>Mitglieder geprüft</span></div>
      <div><strong>{full_universe_count:,}</strong><span>Pairings im vollständigen Universum</span></div>
      <div><strong>{optimization_pool_count}</strong><span>SCIP-Kandidaten im Optimierungspool</span></div>
      <div><strong>{selected_count}</strong><span>kuratierte Karten ausgewählt</span></div>
      <div><strong>{executive_count}</strong><span>Executive-Linien sichtbar</span></div>
      <div><strong>NN</strong><span>Feedbackranker {nn_active}</span></div>
    </section>
    """.replace(",", ".")

    selected_pattern_ids = {clean_display_value(row.get("pattern_id", "")) for _, row in selected.iterrows()}
    alternatives = ranked[
        ~ranked["pattern_id"].astype(str).map(clean_display_value).isin(selected_pattern_ids)
    ].copy()
    alternatives = alternatives[alternatives.get("selected", 0).astype(str) != "1"].copy()
    alt_cards = []
    for _, row in alternatives.iterrows():
        rank = row_rank(row)
        reason = clean_display_value(row.get("why_not_selected", "")) or clean_display_value(row.get("non_selection_reason", ""))
        if not reason:
            reason = "Nicht ausgewählt; als nahe Alternative für Feedback behalten."
        theme = clean_display_value(row.get("convening_theme", ""))
        balance = clean_display_value(row.get("pair_evidence_balance", ""))
        decision = clean_display_value(row.get("action_decision", ""))
        candidate_label = scip_card_label(row, 99)
        candidate_title = card_title_text(row, candidate_label)
        feedback_payload = feedback_payload_for_row(row, rank, False, "feedback_only_candidate")
        payload_attr = html.escape(json.dumps(feedback_payload, ensure_ascii=False), quote=True)
        meta_bits = ["nicht ausgewählt", "Feedback erwünscht"]
        if balance:
            meta_bits.append(f"Balance: {balance}")
        if decision:
            meta_bits.append(f"Decision: {decision}")
        if theme:
            meta_bits.append(theme)
        alt_cards.append(
            f"""
            <article class="candidate-row" data-feedback-payload="{payload_attr}">
              <div class="candidate-main">
                <div class="candidate-title">#{rank} · {html.escape(candidate_title)}</div>
                <div class="candidate-meta">{html.escape(" · ".join(meta_bits))}</div>
                <p>{html.escape(reason)}</p>
                {neural_learning_signal_html(row)}
              </div>
              {feedback_controls}
            </article>
            """
        )
    alternatives_body = "\n".join(alt_cards) or "<p>Keine weiteren Kandidaten zur Bewertung.</p>"

    html_doc = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="utf-8" />
  <title>SCIP Match Selector</title>
  <link rel="icon" type="image/png" href="sor-tab-o.png" />
  <link rel="apple-touch-icon" href="sor-tab-o.png" />
  <style>
    :root {{
      --bg:#05070A;
      --panel:#07111F;
      --card:#F5F7FA;
      --ink:#05070A;
      --muted:#647183;
      --line:#D8E3ED;
      --blue:#00AEEF;
      --blue-soft:#EAF8FF;
      --gold:#F5B82E;
      --gold-soft:#FFF6D9;
      --white:#F5F7FA;
      --shadow:0 18px 44px rgba(0,0,0,.24);
    }}
    * {{ box-sizing:border-box; }}
    body {{
      margin:0;
      background:linear-gradient(180deg,#05070A 0%,#07111F 45%,#091320 100%);
      color:var(--white);
      font-family:Inter,Segoe UI,system-ui,sans-serif;
    }}
    body::before {{
      content:"";
      position:fixed;
      inset:0;
      pointer-events:none;
      opacity:.16;
      background:repeating-linear-gradient(90deg,rgba(0,174,239,.06) 0 1px,transparent 1px 92px),repeating-linear-gradient(0deg,rgba(0,174,239,.045) 0 1px,transparent 1px 92px);
      mask-image:linear-gradient(180deg,black,transparent 78%);
    }}
    main {{ position:relative; max-width:1120px; margin:0 auto; padding:34px 18px 70px; }}
    header, .card, .candidate-row {{ background:var(--card); border:1px solid rgba(216,227,237,.92); border-top:2px solid rgba(0,174,239,.58); border-radius:8px; box-shadow:var(--shadow); }}
    header {{
      position:relative;
      padding:30px;
      margin-bottom:20px;
      overflow:hidden;
      background:linear-gradient(135deg,rgba(7,17,31,.98),rgba(5,7,10,.94));
      color:var(--white);
      border-color:rgba(0,174,239,.28);
    }}
    .selector-brand {{ position:relative; z-index:1; display:flex; align-items:center; gap:20px; min-width:0; }}
    .selector-brand img {{ width:min(260px,30vw); height:auto; aspect-ratio:16 / 9; object-fit:contain; object-position:center; border-radius:8px; border:1px solid rgba(0,174,239,.20); background:#000; box-shadow:0 0 26px rgba(0,174,239,.14); flex:0 0 auto; }}
    h1 {{ margin:0; font-size:clamp(1.65rem,2.6vw,2.35rem); letter-spacing:.04em; text-transform:uppercase; font-weight:850; }}
    @media (max-width:720px) {{ .selector-brand img {{ width:min(260px,72vw); }} }}
    h2.section-title {{ margin:30px 0 8px; font-size:1rem; color:var(--blue); text-transform:uppercase; letter-spacing:.12em; }}
    .section-note {{ color:#AAB3C0; margin:0 0 14px; line-height:1.45; }}
    .meta {{ color:#AAB3C0; margin-top:9px; letter-spacing:.04em; }}
    header p {{ color:#AAB3C0; }}
    .grid {{ display:grid; gap:14px; }}
    .card {{ padding:22px; position:relative; color:var(--ink); }}
    .card h2 {{ margin:0 72px 10px 0; font-size:1.08rem; color:var(--ink); }}
    .card p {{ line-height:1.55; margin:8px 0; }}
    .card-label {{ color:var(--blue); font-weight:900; text-transform:uppercase; letter-spacing:.09em; font-size:.73rem; margin-bottom:6px; }}
    .critical-decision-grid, .design-arc, .decision-triplet {{ border:1px solid #dfe7ef; background:#fff; border-left:3px solid var(--blue); border-radius:8px; padding:11px; margin:12px 0; display:grid; gap:9px; }}
    .compact-card-copy h2 {{ margin:0 0 10px 0; font-size:1.04rem; line-height:1.3; }}
    .decision-summary {{ color:#172033; font-weight:650; line-height:1.5; }}
    .decision-triplet p {{ margin:0; font-size:.9rem; line-height:1.45; }}
    .compact-card h2 {{ margin-right:0; }}
    .compact-card strong {{ color:#172033; }}
    .candidate-row {{ padding:16px 18px; display:grid; gap:10px; color:var(--ink); }}
    .candidate-title {{ color:#172033; font-weight:900; line-height:1.35; }}
    .candidate-meta {{ color:#647183; font-weight:820; font-size:.84rem; margin-top:3px; text-transform:uppercase; letter-spacing:.04em; }}
    .candidate-row p {{ margin:7px 0 0; color:#344257; line-height:1.45; }}
    .learning-signal {{ display:inline-flex; width:max-content; max-width:100%; margin:8px 0 2px; border-radius:999px; padding:6px 10px; font-size:.78rem; font-weight:850; border:1px solid #d6e2ef; color:#263348; background:#f7fafc; }}
    .learning-signal-positiv {{ border-color:rgba(34,197,94,.42); background:rgba(34,197,94,.11); color:#14532d; }}
    .learning-signal-neutral {{ border-color:rgba(245,158,11,.42); background:rgba(245,158,11,.12); color:#78350f; }}
    .learning-signal-negativ {{ border-color:rgba(239,68,68,.38); background:rgba(239,68,68,.10); color:#7f1d1d; }}
    .feedback-box {{ border-top:1px solid #d9e1ea; margin-top:14px; padding-top:12px; display:flex; flex-wrap:wrap; gap:8px; align-items:center; }}
    .candidate-row .feedback-box {{ margin-top:4px; }}
    .feedback-box span {{ color:#647183; font-weight:850; margin-right:2px; }}
    .feedback-box button {{ border:1px solid rgba(0,174,239,.28); background:white; color:#172033; border-radius:999px; padding:7px 11px; font-weight:800; cursor:pointer; }}
    .feedback-box button:hover {{ border-color:var(--blue); background:var(--blue-soft); }}
    .feedback-comment-box {{ flex-basis:100%; margin-top:8px; display:grid; gap:7px; }}
    .feedback-comment-box[hidden] {{ display:none; }}
    .feedback-comment-box label {{ font-weight:800; color:#647183; }}
    .feedback-comment-box select {{ border:1px solid #cfe0f1; border-radius:8px; padding:9px; font:inherit; font-weight:700; color:#172033; }}
    .feedback-comment-box textarea {{ width:100%; box-sizing:border-box; border:1px solid #cfe0f1; border-radius:8px; padding:9px; font:inherit; }}
    .feedback-status {{ color:#647183; flex-basis:100%; }}
    .button-link {{ display:inline-block; margin-top:14px; border:1px solid rgba(0,174,239,.52); background:#07111F; color:var(--white); border-radius:999px; padding:9px 13px; text-decoration:none; font-weight:850; text-transform:uppercase; letter-spacing:.06em; }}
    .button-link:hover {{ background:var(--blue); color:#05070A; }}
    .trust-block {{ display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:10px; margin:0 0 22px; }}
    .trust-block div {{ border:1px solid rgba(0,174,239,.24); background:linear-gradient(135deg,rgba(7,17,31,.94),rgba(5,7,10,.88)); border-radius:8px; padding:12px 13px; min-height:74px; }}
    .trust-block strong {{ display:block; color:var(--blue); font-size:1.25rem; line-height:1.1; margin-bottom:5px; }}
    .trust-block span {{ display:block; color:#AAB3C0; font-size:.78rem; line-height:1.25; font-weight:800; }}
    @media (max-width:980px) {{ .trust-block {{ grid-template-columns:repeat(3,minmax(0,1fr)); }} }}
    @media (max-width:620px) {{ .trust-block {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} }}
    a {{ color:var(--blue); }}
  </style>
</head>
<body>
<main>
  <header>
    <div class="selector-brand">
      <img src="sor-logo-full.png" alt="SOR Strategic Opportunity Radar Logo" />
      <h1>SCIP Match Selector</h1>
    </div>
    <div class="meta">Kuratierte Empfehlungen plus Feedback-Leaderboard ? SCIP-basierte Auswahl aus aktuellen Mitgliedssignalen ? erstellt {datetime.now().strftime('%d.%m.%Y, %H:%M:%S')}</div>
    <p>Profile: <a href="{profile_path.name}">{profile_path.name}</a> ? Patterns: <a href="{pattern_path.name}">{pattern_path.name}</a></p>
    {dashboard_link}
  </header>
  {trust_block}

  <h2 class="section-title">Executive-Linien</h2>
  <p class="section-note">Die ersten sechs Karten sind für die schnelle strategische Sichtung kuratiert: klare Akteurslogik, sichtbarer Anlass, nächster prüfbarer Schritt.</p>
  <section class="grid">
    {selected_body}
  </section>

  <h2 class="section-title">Weitere kuratierte Karten</h2>
  <p class="section-note">Diese vier Karten bleiben im Portfolio, eignen sich aber eher für Validierung, Watchlist oder Feedback als für die erste Leitungssicht.</p>
  <section class="grid">
    {additional_body}
  </section>

  <h2 class="section-title">Nicht ausgewählte Kandidaten für Feedback</h2>
  <p class="section-note">Diese Kandidaten sind keine Empfehlungen, bleiben aber sichtbar, damit SCIP aus nahen Alternativen und verworfenen Matchings lernen kann.</p>
  <section class="grid">
    {alternatives_body}
  </section>
</main>
<script>
  function defaultFeedbackTarget(label) {{
    if (label === 'useful') return 'actor_match';
    if (label === 'good_topic_wrong_actors' || label === 'wrong_connection') return 'actor_match';
    if (label === 'not_relevant') return 'rlp_relevance';
    return '';
  }}
  function defaultFeedbackReason(label) {{
    if (label === 'useful') return 'good_match';
    if (label === 'not_relevant') return 'missing_rlp_application_anchor';
    return '';
  }}
  function prepareScipFeedback(button, label) {{
    const card = button.closest('[data-feedback-payload]');
    const status = card.querySelector('.feedback-status');
    const box = card.querySelector('.feedback-comment-box');
    const textarea = box.querySelector('textarea');
    const targetSelect = box.querySelector('.feedback-target-select');
    const reasonSelect = box.querySelector('.feedback-reason-select');
    card.dataset.pendingFeedbackLabel = label;
    box.hidden = false;
    if (targetSelect && !targetSelect.value) targetSelect.value = defaultFeedbackTarget(label);
    if (reasonSelect && !reasonSelect.value) reasonSelect.value = defaultFeedbackReason(label);
    textarea.focus();
    status.textContent = 'Feedback kann mehrfach pro Karte gespeichert werden';
  }}
  function cancelScipFeedback(button) {{
    const card = button.closest('[data-feedback-payload]');
    const status = card.querySelector('.feedback-status');
    const box = card.querySelector('.feedback-comment-box');
    const textarea = box.querySelector('textarea');
    const selects = box.querySelectorAll('select');
    card.dataset.pendingFeedbackLabel = '';
    textarea.value = '';
    selects.forEach((select) => select.value = '');
    box.hidden = true;
    status.textContent = '';
  }}
  async function submitScipFeedback(button) {{
    const card = button.closest('[data-feedback-payload]');
    const status = card.querySelector('.feedback-status');
    const box = card.querySelector('.feedback-comment-box');
    const textarea = box.querySelector('textarea');
    const targetSelect = box.querySelector('.feedback-target-select');
    const reasonSelect = box.querySelector('.feedback-reason-select');
    const label = card.dataset.pendingFeedbackLabel || '';
    if (!label) return;
    let payload = {{}};
    try {{ payload = JSON.parse(card.dataset.feedbackPayload || '{{}}'); }} catch (err) {{ payload = {{}}; }}
    payload.label = label;
    payload.human_feedback_label = label;
    payload.feedback_target = targetSelect ? targetSelect.value : '';
    payload.feedback_dimension = payload.feedback_target;
    payload.reason_category = reasonSelect ? reasonSelect.value : '';
    payload.feedback_reason_category = payload.reason_category;
    payload.human_comment = textarea.value.trim();
    payload.human_feedback_comment = payload.human_comment;
    payload.timestamp = new Date().toISOString();
    payload.created_at = payload.timestamp;
    status.textContent = 'speichere Feedback ...';
    try {{
      const response = await fetch('http://127.0.0.1:8766/feedback/scip', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify(payload)
      }});
      if (!response.ok) throw new Error('feedback server unavailable');
      status.textContent = 'Feedback gespeichert';
    }} catch (err) {{
      const key = 'sor_scip_feedback_queue';
      const queue = JSON.parse(localStorage.getItem(key) || '[]');
      queue.push(payload);
      localStorage.setItem(key, JSON.stringify(queue));
      status.textContent = 'Feedback lokal vorgemerkt';
    }}
    textarea.value = '';
    if (targetSelect) targetSelect.value = '';
    if (reasonSelect) reasonSelect.value = '';
    box.hidden = true;
    card.dataset.pendingFeedbackLabel = '';
  }}
</script>
</body>
</html>"""
    html_path.write_text(repair_mojibake(html_doc), encoding="utf-8")


def clean_display_value(value: Any) -> str:
    try:
        if value is None or pd.isna(value):
            return ""
    except Exception:
        if value is None:
            return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    if text.lower() in {"nan", "none", "null", "nat"}:
        return ""
    return repair_mojibake(text)


def repair_mojibake(text: str) -> str:
    if not text:
        return ""
    repaired = str(text)
    replacements = {
        chr(0x00c3) + chr(0x0192) + chr(0x00a4): "\u00e4",
        chr(0x00c3) + chr(0x0192) + chr(0x00b6): "\u00f6",
        chr(0x00c3) + chr(0x0192) + chr(0x00bc): "\u00fc",
        chr(0x00c3) + chr(0x0192) + chr(0x009f): "\u00df",
        chr(0x00c3) + chr(0x201a) + chr(0x00b7): "\u00b7",
        chr(0x00c3) + chr(0x201a): "",
        "\u00c3\u00a4": "\u00e4",
        "\u00c3\u00b6": "\u00f6",
        "\u00c3\u00bc": "\u00fc",
        "\u00c3\u009f": "\u00df",
        "\u00c3\u201e": "\u00c4",
        "\u00c3\u2013": "\u00d6",
        "Ö": "\u00d6",
        "LÖWEN": "L\u00d6WEN",
        "LÖWEN": "L\u00d6WEN",
        "LÖWEN": "L\u00d6WEN",
        "L\\u00c3\\u0096WEN": "L\\u00d6WEN",
            "\u00c3\u0153": "\u00dc",
            "\u00c2\u00b7": "\u00b7",
            "\u00c2": "",
            "\u00e2\u20ac\u201c": "\u2013",
            "\u00e2\u20ac\u201d": "\u2014",
            "\u00e2\u20ac\u017e": "\u201e",
            "\u00e2\u20ac\u0153": "\u201c",
            "\u00e2\u20ac\u009d": "\u201d",
            "\u00e2\u20ac\u2122": "\u2019",
            "–": "\u2013",
            "–": "\u2013",
            "—": "\u2014",
            "\u00c3\u00a2\u00e2\u201a\u00ac\u00e2\u20ac\u0153": "\u2013",
            "\u00c3\u00a2\u00e2\u201a\u00ac\u00c2\u009d": "\u201d",
        }
    for bad, good in replacements.items():
        repaired = repaired.replace(bad, good)
    for _ in range(2):
        if not any(marker in repaired for marker in ("\u00c3", "\u00c2", "\u00e2", "\u0192")):
            break
        try:
            fixed = repaired.encode("cp1252").decode("utf-8")
        except UnicodeError:
            break
        if fixed == repaired:
            break
        repaired = fixed
    return repaired
def refresh_main_dashboard() -> Optional[Path]:
    dashboard_path = REPORT_DIR / "zirp_dashboard.html"
    history_path = REPORT_DIR / "zirp_dashboard_history.json"
    weekly_script = next(
        (
            candidate
            for candidate in [
                PROJECT_DIR / "prototype_v2_final.py",
                PROJECT_DIR / "prototype_v2.py",
            ]
            if candidate.exists()
        ),
        PROJECT_DIR / "prototype_v2_final.py",
    )
    if not dashboard_path.exists() or not history_path.exists() or not weekly_script.exists():
        return None

    spec = importlib.util.spec_from_file_location("zirp_weekly_dashboard", weekly_script)
    if spec is None or spec.loader is None:
        return None

    try:
        os.environ["STATIC_ZIRP_EXCEL_PATH"] = str(STATIC_CANDIDATE_PATH)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        analyzer = module.ZIRPWeeklyAnalyzer()
        history = analyzer.load_dashboard_history(history_path)
        dashboard_path.write_text(analyzer.render_dashboard_html(history), encoding="utf-8")
        return dashboard_path
    except Exception as exc:
        print(f"Dashboard refresh: übersprungen ({exc})")
        return None


def main() -> None:
    events_path = latest_file("zirp_ereignisse")
    members_path = latest_file("zirp_mitglied_homepages")
    weighted_path = latest_file("zirp_begriffe_gewichtet")

    print("Using events:", events_path)
    print("Using members:", members_path)
    print("Using weighted terms:", weighted_path)

    members_df = read_csv(members_path)
    events_df_raw = read_csv(events_path)
    events_df = filter_events_to_period(events_df_raw)
    weighted_df = read_csv(weighted_path)
    dynamic_terms = load_dynamic_filter_terms(weighted_path, events_df)
    active_dynamic_count = sum(len(values) for values in dynamic_terms.values())
    print(
        "Dynamic filter terms: "
        f"problem={len(dynamic_terms.get('problem', set()))}, "
        f"action={len(dynamic_terms.get('action', set()))}, "
        f"asset={len(dynamic_terms.get('asset', set()))}, "
        f"noise={len(dynamic_terms.get('noise', set()))}, "
        f"active={active_dynamic_count}"
    )
    print(f"Dynamic term memory: {DYNAMIC_TERM_MEMORY_PATH}")

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    profiles = build_member_profiles(members_df, events_df)

    optimizer_workbook_path = STATIC_OPTIMIZER_PATH if STATIC_OPTIMIZER_PATH.exists() else STATIC_CANDIDATE_PATH
    if USE_STATIC_CANDIDATES and optimizer_workbook_path.exists():
        refresh_static_scip_optimization_pool(optimizer_workbook_path, events_df)
        patterns = load_static_candidate_patterns_from_excel(optimizer_workbook_path, events_df)
        print(f"Static Excel candidate universe P loaded: {len(patterns)} patterns from {optimizer_workbook_path}")
    else:
        if USE_STATIC_CANDIDATES:
            print(f"Static candidate workbook missing; falling back to generated patterns: {optimizer_workbook_path}")
        patterns = generate_meeting_patterns(profiles)
        patterns = filter_pair_only_patterns(patterns)
        print(f"Candidate pair patterns after filters: {len(patterns)}")

    active_members = sorted({
        member
        for pattern in patterns
        for member in pattern.members
    })

    output_dir = REPORT_DIR
    profile_path = output_dir / f"zirp_member_topic_profiles_{timestamp}.csv"
    pattern_path = output_dir / f"zirp_meeting_patterns_{timestamp}.csv"
    rulebased_pattern_path = output_dir / f"zirp_meeting_patterns_rulebased_{timestamp}.csv"
    ai_pattern_path = output_dir / f"zirp_meeting_patterns_ai_{timestamp}.csv"
    leaderboard_path = output_dir / f"zirp_match_leaderboard_{timestamp}.csv"
    incidence_matrix_path = output_dir / f"zirp_scip_incidence_matrix_{timestamp}.csv"
    model_summary_path = output_dir / f"zirp_match_model_summary_{timestamp}.json"
    dynamic_terms_path = output_dir / f"zirp_dynamic_term_candidates_{timestamp}.csv"
    feedback_summary_path = output_dir / f"zirp_feedback_learning_summary_{timestamp}.json"
    lp_path = output_dir / f"zirp_meeting_optimizer_{timestamp}.lp"
    scip_log_path = output_dir / f"zirp_meeting_optimizer_scip_{timestamp}.log"
    html_path = output_dir / f"zirp_meeting_recommendations_{timestamp}.html"
    ai_html_path = output_dir / f"zirp_meeting_recommendations_ai_{timestamp}.html"

    export_profiles(profiles, profile_path)
    export_scip_incidence_matrix(patterns, incidence_matrix_path)
    write_scip_model(patterns, active_members, lp_path)

    try:
        selected_ids, scip_output = run_scip(lp_path)
        if not selected_ids:
            selected_ids = greedy_fallback(patterns)
            scip_output += "\n\nNo SCIP solution parsed; used greedy fallback.\n"
    except Exception as exc:
        selected_ids = greedy_fallback(patterns)
        scip_output = f"SCIP failed, used greedy fallback: {exc}\n"

    scip_log_path.write_text(scip_output, encoding="utf-8", errors="replace")
    write_model_summary(
        model_summary_path,
        timestamp=timestamp,
        patterns=patterns,
        selected_ids=selected_ids,
        solver_output=scip_output,
    )
    patterns_df = export_patterns(patterns, selected_ids, pattern_path)
    patterns_df = apply_convening_selection_filters(patterns_df)
    patterns_df.to_csv(rulebased_pattern_path, index=False, encoding="utf-8-sig")
    if USE_OLLAMA_FOR_CONVENING_TEXT:
        ai_patterns_df, ai_updated = enrich_selected_patterns_with_ollama(patterns_df)
        ai_patterns_df = apply_convening_selection_filters(ai_patterns_df)
        ai_patterns_df.to_csv(ai_pattern_path, index=False, encoding="utf-8-sig")
        if ai_updated:
            patterns_df = ai_patterns_df
    else:
        ai_patterns_df = patterns_df
        ai_updated = 0
        print("KI Convening Text: uebersprungen (ZIRP_CONVENING_AI=0)")
    patterns_df.to_csv(pattern_path, index=False, encoding="utf-8-sig")
    update_opportunity_history(
        patterns_df,
        source_name="ai" if USE_OLLAMA_FOR_CONVENING_TEXT and ai_updated else "rulebased",
        timestamp=timestamp,
    )
    patterns_df = attach_history_columns(patterns_df)
    ai_patterns_df = attach_history_columns(ai_patterns_df)
    patterns_df.to_csv(pattern_path, index=False, encoding="utf-8-sig")
    export_selection_leaderboard(patterns_df, leaderboard_path)
    export_dynamic_term_candidates(dynamic_terms_path, patterns_df)
    summarize_feedback_learning(patterns_df, feedback_summary_path)
    if USE_OLLAMA_FOR_CONVENING_TEXT:
        ai_patterns_df.to_csv(ai_pattern_path, index=False, encoding="utf-8-sig")
    audit_path = output_dir / f"zirp_convening_dominance_audit_{timestamp}.csv"
    audit_institutional_dominance(patterns_df, audit_path)
    render_html(patterns_df, profile_path, pattern_path, html_path)
    if USE_OLLAMA_FOR_CONVENING_TEXT:
        render_html(ai_patterns_df, profile_path, ai_pattern_path, ai_html_path)
    latest_html_path = output_dir / "zirp_meeting_recommendations_latest.html"
    latest_html_path.write_text(html_path.read_text(encoding="utf-8"), encoding="utf-8")
    latest_ai_html_path = None
    if USE_OLLAMA_FOR_CONVENING_TEXT:
        latest_ai_html_path = output_dir / "zirp_meeting_recommendations_ai_latest.html"
        latest_ai_html_path.write_text(ai_html_path.read_text(encoding="utf-8"), encoding="utf-8")
    dashboard_path = refresh_main_dashboard()

    print("\nCreated:")
    print("  Profiles:", profile_path)
    print("  Patterns:", pattern_path)
    print("  Rule-based Patterns:", rulebased_pattern_path)
    print("  Match Leaderboard:", leaderboard_path)
    print("  SCIP Incidence Matrix:", incidence_matrix_path)
    print("  Dynamic Terms:", dynamic_terms_path)
    print("  Dynamic Term Memory:", DYNAMIC_TERM_MEMORY_PATH)
    print("  Feedback Summary:", feedback_summary_path)
    print("  Model Summary:", model_summary_path)
    if USE_OLLAMA_FOR_CONVENING_TEXT:
        print("  AI Patterns:", ai_pattern_path)
    print("  Dominance audit:", audit_path)
    print("  SCIP LP:", lp_path)
    print("  SCIP log:", scip_log_path)
    print("  HTML:", html_path)
    if USE_OLLAMA_FOR_CONVENING_TEXT:
        print("  AI HTML:", ai_html_path)
    print("  Latest HTML:", latest_html_path)
    if latest_ai_html_path:
        print("  Latest AI HTML:", latest_ai_html_path)
    if dashboard_path:
        print("  Dashboard:", dashboard_path)
    print("\nSelected matches:")
    for _, row in patterns_df[patterns_df["selected"] == 1].iterrows():
        members = repair_mojibake(str(row.get("members", "")))
        print(f"  - {members}  score={row['score']}")


if __name__ == "__main__":
    main()

