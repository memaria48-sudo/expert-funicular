# Strategic Opportunity Radar

SOR detects weak signals, scores relevance, proposes convening opportunities,
and writes a weekly briefing/dashboard.

## Current Safe Architecture

The working prototype is frozen in `prototype_v1.py`.

`one more try.py` remains the PyCharm-compatible launcher and now delegates to:

```text
sor/app/main.py
```

This gives the project a clean structure without breaking the current crawler,
SCIP optimizer, cache, Groq/OpenAI reporting, or GitHub Pages dashboard.

## Folder Jobs

```text
sor/
  app/
    main.py              # main application entry point
    config.py            # runtime/config loading helpers
    models.py            # SignalEvent, MatchCandidate, FeedbackRecord
    legacy.py            # adapter around frozen prototype_v1.py

    crawler/             # discovery, fetching, extraction, cache
    scoring/             # features, rules, profiles, SCIP/matching boundary
    llm/                 # signal judge and report writer boundaries
    reports/             # briefing/dashboard/export/output manifest
    nn/                  # future adaptive reranker, not active yet

configs/                 # domain YAML configs
data/                    # raw, processed, feedback data
outputs/                 # future generated outputs
tests/                   # structure/import tests
```

## Run

Existing PyCharm command still works:

```powershell
python "one more try.py"
```

New structured entry point also works:

```powershell
python -m sor.app.main
python -m sor.app.main --describe
python -m sor.app.main --manifest
```

Run the isolated SCIP + ML embedding demo:

```powershell
python -m sor.app.main --nn-scip-demo
```

This demo follows the PySCIPOpt-ML basic example: train two predictors on
synthetic nonlinear functions, embed them as SCIP constraints, optimize, and
check the embedding error. It is not yet part of the production SOR scoring
pipeline.

Run the local SCIP feedback collector before reviewing the dashboard:

```powershell
python -m sor.app.main --feedback-server
```

Then open `zirp_berichte/scip_archive.html` and use the feedback buttons on
each SCIP card. Feedback is saved to:

```text
data/feedback/scip_feedback.jsonl
```

The optimizer reads that file on later runs and applies a small learned
adjustment to SCIP pattern scores. If the feedback server is not running, the
dashboard keeps feedback in browser local storage as a fallback.

## NN Plan

The NN should not replace SCIP. It should become a learned adjustment layer:

```text
rule-based SCIP score
+ organization priority score
+ actor-topic match score
+ LLM judgment
+ NN learned adjustment
= final recommendation score
```

Before training, collect feedback in `data/feedback/`.

The first production NN step should be:

1. collect dashboard feedback,
2. convert `SignalEvent` and `MatchCandidate` objects into feature rows,
3. train a small reranker,
4. embed or apply that reranker as a score adjustment after SCIP/rule scoring.
