from __future__ import annotations

import csv
import html
import json
import re
from pathlib import Path
from typing import Any


PROJECT_DIR = Path(__file__).resolve().parent
REPORTS = PROJECT_DIR / "zirp_berichte"


def h(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def repair_mojibake(value: object) -> str:
    text = str(value or "")
    if not text:
        return ""
    replacements = {
        "Ã¤": "ä", "Ã¶": "ö", "Ã¼": "ü", "ÃŸ": "ß",
        "Ã„": "Ä", "Ã–": "Ö", "Ãœ": "Ü",
        "Â·": "·", "Â©": "©", "Â": "",
        "â€“": "–", "â€”": "—", "â€ž": "„", "â€œ": "“",
        "â€": "”", "â€™": "’", "â€˜": "‘", "â€¦": "…",
        "Ã¢â‚¬â€œ": "–", "Ã¢â‚¬â€": "—", "Ã¢â‚¬Å¾": "„",
        "Ã¢â‚¬Å“": "“", "Ã¢â‚¬Â": "”", "Ã¢â‚¬â„¢": "’",
        "ÃƒÂ¤": "ä", "ÃƒÂ¶": "ö", "ÃƒÂ¼": "ü", "ÃƒÅ¸": "ß",
        "Fachkraefte": "Fachkräfte", "Klaerung": "Klärung",
        "klaeren": "klären", "pruefen": "prüfen", "fuer": "für",
        "waehrend": "während", "koennen": "können", "ueber": "über",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    for _ in range(2):
        if not any(marker in text for marker in ("Ã", "Â", "â", "ƒ")):
            break
        try:
            fixed = text.encode("cp1252").decode("utf-8")
        except UnicodeError:
            break
        if fixed == text:
            break
        text = fixed
    return text


def latest_file(pattern: str) -> Path | None:
    files = sorted(REPORTS.glob(pattern), key=lambda path: path.stat().st_mtime, reverse=True)
    return files[0] if files else None


def load_current_history() -> dict[str, Any]:
    path = REPORTS / "zirp_dashboard_history.json"
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if isinstance(data, list) and data:
        return data[0] if isinstance(data[0], dict) else {}
    return {}


def clean_text(value: object) -> str:
    return re.sub(r"\s+", " ", repair_mojibake(value)).strip()


def clean_url(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    replacements = {
        "Ã¤": "ä", "Ã¶": "ö", "Ã¼": "ü", "ÃŸ": "ß",
        "Ã„": "Ä", "Ã–": "Ö", "Ãœ": "Ü",
        "Â": "", "â€“": "–", "â€”": "—",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text


def paragraph_html(text: str) -> str:
    fixed = repair_mojibake(text)
    fixed = re.sub(r"^Wochenanalyse\s*", "", fixed.strip(), flags=re.I)
    paragraphs = [re.sub(r"\s+", " ", p).strip() for p in re.split(r"\n{2,}", fixed) if p.strip()]
    if not paragraphs:
        return "<p>Noch keine Wochenanalyse gespeichert.</p>"
    title = ""
    if paragraphs and ":" not in paragraphs[0][:40] and len(paragraphs[0]) < 120:
        title = f"<h1>{h(paragraphs.pop(0))}</h1>"
    body = "\n".join(f"<p>{h(p)}</p>" for p in paragraphs)
    return title + body


COMMON_CSS = """
:root {
  --bg: #03070d; --panel: rgba(4,16,30,.86); --line: rgba(78,174,255,.34);
  --blue: #1f9bff; --blue2:#42c8ff; --ink:#08101d; --white:#f7fbff;
  --green:#38c558; --yellow:#ffc83d; --red:#ff3d3d;
}
*{box-sizing:border-box} html{background:var(--bg)}
body{margin:0;min-height:100vh;color:var(--white);font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:radial-gradient(circle at 16% 10%,rgba(31,155,255,.22),transparent 32rem),radial-gradient(circle at 84% 20%,rgba(22,112,214,.16),transparent 30rem),linear-gradient(180deg,#02050a 0%,#06101d 58%,#03070d 100%)}
body:before{content:"";position:fixed;inset:0;pointer-events:none;opacity:.16;background:linear-gradient(90deg,rgba(72,159,244,.08) 1px,transparent 1px),linear-gradient(0deg,rgba(72,159,244,.06) 1px,transparent 1px);background-size:86px 86px;mask-image:linear-gradient(180deg,#000 0%,transparent 72%)}
a{color:inherit}.shell{width:min(1720px,calc(100% - 48px));margin:0 auto;position:relative;z-index:1}
.topbar{margin:18px auto 16px;padding:20px 26px;min-height:118px;display:flex;align-items:center;justify-content:space-between;gap:24px;border:1px solid rgba(50,141,225,.32);border-radius:8px;background:linear-gradient(135deg,rgba(3,15,29,.9),rgba(3,10,20,.78));box-shadow:0 22px 50px rgba(0,0,0,.28),inset 0 1px 0 rgba(255,255,255,.04)}
.brand{display:inline-flex;align-items:center;text-decoration:none}.brand img{width:260px;height:auto;display:block}
.nav{display:flex;align-items:center;justify-content:flex-end;gap:16px;flex-wrap:wrap}.nav a{min-width:144px;text-align:center;text-decoration:none;text-transform:uppercase;letter-spacing:.06em;font-size:14px;font-weight:800;color:#dfe8f4;border:1px solid rgba(173,198,224,.26);border-radius:999px;padding:12px 24px;background:rgba(255,255,255,.02)}.nav a.active{color:#fff;background:linear-gradient(135deg,#105ee9,#2fc7ff);border-color:rgba(65,196,255,.65);box-shadow:0 0 28px rgba(31,155,255,.42)}
.meta-strip{display:flex;flex-wrap:wrap;gap:36px;align-items:center;margin-bottom:20px;padding:13px 32px;border:1px solid rgba(50,141,225,.34);border-radius:8px;background:rgba(4,16,30,.7);color:#d7deea;font-size:15px}.meta-item{display:inline-flex;align-items:center;gap:14px}.meta-icon{width:22px;height:22px;display:inline-grid;place-items:center;color:var(--blue2)}.meta-dot{color:rgba(194,208,224,.38);font-weight:800}
.panel{border:1px solid rgba(50,141,225,.34);border-radius:8px;background:radial-gradient(circle at 10% 20%,rgba(31,155,255,.16),transparent 28rem),linear-gradient(135deg,rgba(4,17,32,.9),rgba(3,10,19,.86));box-shadow:0 24px 70px rgba(0,0,0,.34),inset 0 1px 0 rgba(255,255,255,.04)}
.footer{margin:26px auto 16px;padding:16px 20px;display:grid;grid-template-columns:1fr 2fr 1fr;gap:20px;align-items:center;border-top:1px solid rgba(50,141,225,.22);color:#c7d0dc}.footer strong{color:var(--blue2)}.footer-center{text-align:center;color:#aeb8c7}.footer-right{text-align:right}.status{display:inline-flex;align-items:center;gap:12px}.status-dot{width:18px;height:18px;border-radius:50%;background:var(--green);box-shadow:0 0 22px rgba(56,197,88,.45)}
@media(max-width:900px){.shell{width:min(100% - 28px,1540px)}.topbar{align-items:flex-start;flex-direction:column}.nav{width:100%;justify-content:flex-start}.nav a{min-width:auto;flex:1 1 160px}.meta-strip{flex-direction:column;align-items:flex-start;gap:12px;padding:14px 18px}.meta-dot{display:none}.footer{grid-template-columns:1fr}.footer-center,.footer-right{text-align:left}}
"""


HOME_CSS = COMMON_CSS + """
.home-grid{display:grid;grid-template-columns:minmax(0,2.45fr) minmax(310px,.8fr);gap:18px}.article{background:#fbfcfe;color:var(--ink);border:1px solid rgba(91,180,255,.46);border-radius:8px;padding:40px 42px;box-shadow:0 22px 44px rgba(0,0,0,.28)}.article h1{margin:0 0 18px;color:#07101f;text-transform:uppercase;letter-spacing:.02em;font-size:clamp(32px,3.25vw,46px);line-height:1.12}.article p{margin:0 0 17px;font-size:19px;line-height:1.58;color:#111827}.article p:first-of-type{padding:18px 22px;border:1px solid rgba(31,155,255,.55);border-left:5px solid var(--blue);border-radius:5px;color:#0755b5;font-size:19px;line-height:1.58;font-weight:800;background:rgba(238,248,255,.72)}.sources{padding:30px 34px}.sources h2,.mini-panel h2{margin:0 0 20px;text-transform:uppercase;letter-spacing:.06em;font-size:20px}.source-item{padding:20px 0;border-top:1px solid rgba(166,205,240,.18)}.source-item:first-of-type{border-top:1px solid rgba(166,205,240,.28)}.source-title{display:block;color:#f6fbff;font-size:17px;line-height:1.32;font-weight:800;text-decoration:none}.source-title:hover{color:var(--blue2);text-decoration:underline;text-underline-offset:4px}.source-member{margin-top:10px;color:#57b8ff;font-size:15px}.insight-grid{display:grid;grid-template-columns:1fr .82fr 1.05fr;gap:16px;margin-top:16px}.mini-panel{padding:22px 24px}.term-cloud{display:flex;flex-wrap:wrap;gap:9px 11px}.term{display:inline-flex;min-width:112px;justify-content:center;gap:6px;padding:5px 12px;border:1px solid rgba(31,155,255,.6);border-radius:8px;color:#dceefe;background:rgba(5,39,77,.72);font-size:13px}.bar-row{margin:14px 0 18px;color:#dce6f1}.bar-label{display:flex;justify-content:space-between;gap:14px;font-size:14px}.bar-track{margin-top:8px;height:8px;border-radius:99px;background:rgba(89,159,219,.18);overflow:hidden}.bar-fill{height:100%;border-radius:inherit;background:linear-gradient(90deg,#1169f7,#42c8ff)}.members-text{color:#d8e3ee;font-size:16px;line-height:1.6}.member-count{display:inline-flex;margin-top:22px;min-width:220px;justify-content:center;padding:9px 14px;border-radius:6px;border:1px solid rgba(31,155,255,.62);color:#58bfff;background:rgba(4,31,58,.8)}@media(max-width:1100px){.home-grid,.insight-grid{grid-template-columns:1fr}}
"""


MATCH_CSS = COMMON_CSS + """
.match-panel{padding:28px 30px 26px}.section-title{margin:0;text-transform:uppercase;letter-spacing:.06em;font-size:24px}.title-row{display:flex;align-items:center;justify-content:space-between;gap:20px;margin-bottom:24px}.title-left{display:flex;align-items:center;gap:18px}.target-icon{width:28px;height:28px;color:var(--blue2)}.button{text-decoration:none;text-transform:uppercase;letter-spacing:.08em;font-size:13px;font-weight:900;padding:13px 28px;border-radius:12px;border:1px solid var(--blue);background:rgba(5,19,35,.86);color:#fff}.match-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:20px}.match-card{min-height:260px;padding:28px 30px 24px;border:1px solid rgba(176,198,221,.32);border-radius:8px;background:radial-gradient(circle at 12% 18%,rgba(51,167,255,.24),transparent 22rem),linear-gradient(135deg,rgba(7,24,43,.98),rgba(4,13,24,.94))}.match-card.top{border-color:rgba(31,155,255,.95);box-shadow:0 0 0 1px rgba(31,155,255,.28),0 0 34px rgba(31,155,255,.12)}.match-head{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:18px;align-items:start}.rank{color:var(--blue2);font-size:14px;font-weight:900;letter-spacing:.06em;text-transform:uppercase}.pair{margin:9px 0 12px;color:#f6fbff;font-size:21px;line-height:1.25;font-weight:850}.chips{display:flex;align-items:center;flex-wrap:wrap;gap:10px}.chip{display:inline-flex;align-items:center;border:1px solid rgba(68,181,255,.6);border-radius:999px;padding:5px 14px;color:#eaf6ff;background:rgba(6,21,38,.82);font-weight:800}.chip.stage{color:#43bbff;font-size:13px;font-weight:750}.fit-dots{display:flex;gap:28px;justify-content:flex-end}.fit-dot{text-align:center;color:#d7e2ef;font-size:12px;font-weight:700}.dot{display:block;width:18px;height:18px;margin:0 auto 9px;border-radius:50%;box-shadow:0 0 18px currentColor;border:1px solid rgba(255,255,255,.3)}.dot.green{color:var(--green);background:var(--green)}.dot.yellow{color:var(--yellow);background:var(--yellow)}.dot.red{color:var(--red);background:var(--red)}.summary{margin:24px 0 0;color:#e1eaf5;font-size:16.5px;line-height:1.55}.next{margin-top:16px;padding-top:15px;border-top:1px solid rgba(167,205,239,.18);color:#f0f6fd;font-size:15px;line-height:1.45}.next strong{color:#fff}@media(max-width:1050px){.match-grid{grid-template-columns:1fr}.match-head{grid-template-columns:1fr}.fit-dots{justify-content:flex-start}}
"""


def nav(active: str, prefix: str) -> str:
    items = [("home", "Wochenanalyse", f"{prefix}zirp_dashboard.html"), ("scip", "SCIP Matchings", f"{prefix}scip_archive.html"), ("archive", "Archiv", f"{prefix}archive.html")]
    return "\n".join(f'<a class="{"active" if key == active else ""}" href="{h(url)}">{h(label)}</a>' for key, label, url in items)


def header(active: str, prefix: str) -> str:
    return f'<header class="topbar"><a class="brand" href="{h(prefix)}zirp_dashboard.html"><img src="{h(prefix)}sor-logo-full.png" alt="SOR Strategic Opportunity Radar Logo"></a><nav class="nav">{nav(active, prefix)}</nav></header>'


def footer(status: str) -> str:
    return f'<footer class="footer"><div><strong>SOR</strong> · Strategic Opportunity Radar</div><div class="footer-center">© 2026 SOR / SCIP method, scoring logic and convening radar concept.<br>All rights reserved.</div><div class="footer-right"><strong style="color:#fff">Version 2.0</strong><br><span class="status">{h(status)} <span class="status-dot"></span></span></div></footer>'


def home_page(prefix: str) -> str:
    current = load_current_history()
    meta = clean_text(current.get("meta_line")) or "102 Mitglieder geprüft · Detect signals. Connect actors. Act in time."
    analysis = paragraph_html(str(current.get("analysis") or ""))
    sources = []
    for item in current.get("top_events", [])[:8]:
        if not isinstance(item, dict):
            continue
        title = clean_text(item.get("title"))
        url = clean_url(item.get("url"))
        member = clean_text(item.get("member"))
        if title and url:
            sources.append(f'<div class="source-item"><a class="source-title" href="{h(url)}" target="_blank" rel="noopener noreferrer">{h(title)}</a><div class="source-member">{h(member)}</div></div>')
    terms = []
    for item in current.get("top_terms", [])[:12]:
        if isinstance(item, dict):
            term = clean_text(item.get("term") or item.get("word") or item.get("name"))
            score = item.get("score") or item.get("count") or item.get("weight") or ""
            if term:
                terms.append(f'<span class="term">{h(term)} · {h(score)}</span>')
    bars = []
    sections = [x for x in current.get("top_sections", []) if isinstance(x, dict)]
    max_count = max([int(x.get("count") or 0) for x in sections] or [1])
    for item in sections:
        label = clean_text(item.get("title") or item.get("key"))
        count = int(item.get("count") or 0)
        width = max(12, int(count / max_count * 100))
        bars.append(f'<div class="bar-row"><div class="bar-label"><span>{h(label)}</span><strong>{count}</strong></div><div class="bar-track"><div class="bar-fill" style="width:{width}%"></div></div></div>')
    members = ", ".join(clean_text(x) for x in current.get("members", []) if clean_text(x))
    return f"""<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>SOR Wochenradar · Strategic Opportunity Radar</title><link rel="icon" type="image/png" href="{prefix}sor-tab-o.png"><link rel="apple-touch-icon" href="{prefix}sor-tab-o.png"><style>{HOME_CSS}</style></head><body><div class="shell">{header("home", prefix)}<section class="meta-strip"><div class="meta-item">{h(meta)} · Detect signals. Connect actors. Act in time.</div></section><main><section class="home-grid"><article class="article">{analysis}</article><aside class="panel sources"><h2>Quellen</h2>{''.join(sources) or '<p>Keine Quellen im Zeitraum.</p>'}</aside></section><section class="insight-grid"><div class="panel mini-panel"><h2>Begriffe</h2><div class="term-cloud">{''.join(terms)}</div></div><div class="panel mini-panel"><h2>Schwerpunkte</h2>{''.join(bars)}</div><div class="panel mini-panel"><h2>Mitglieder</h2><div class="members-text">{h(members)}</div><div class="member-count">{h(current.get("member_count") or 0)} Mitglieder geprüft</div></div></section></main>{footer("System Status: Online")}</div></body></html>"""


def read_matches() -> list[dict[str, str]]:
    path = latest_file("zirp_meeting_patterns_ai_*.csv") or latest_file("zirp_meeting_patterns_*.csv")
    if not path:
        return []
    rows: list[dict[str, str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if str(row.get("selected", "")).strip() == "1":
                rows.append(row)
    rows.sort(key=lambda r: float(r.get("display_points") or 0), reverse=True)
    return rows[:6]


def dot_classes(row: dict[str, str]) -> tuple[str, str, str]:
    pts = float(row.get("display_points") or 0)
    evidence = row.get("live_evidence_status", "")
    action = (row.get("aktionsreife") or "").lower()
    return ("green" if pts >= 27 else "yellow", "green" if "one_sided" in evidence or "balanced" in evidence or "paired" in evidence else "yellow", "green" if "hoch" in action else ("yellow" if pts >= 29 else "red"))


def stage_label(row: dict[str, str], index: int) -> str:
    if index == 0:
        return "Validierungslinie"
    raw = clean_text(row.get("public_status") or row.get("convening_stage"))
    return {"Hauptlinie": "Validierungslinie", "Beobachten": "Beobachtungslinie"}.get(raw, raw or "Sondierungslinie")


def match_card(row: dict[str, str], index: int) -> str:
    pair = clean_text(row.get("members") or row.get("member_pair_key")).replace("|", "×")
    score = int(round(float(row.get("display_points") or 0)))
    summary = clean_text(row.get("decision_summary") or row.get("reason"))
    next_line = clean_text(row.get("next_line") or row.get("next_best_action") or "Nächsten validierbaren Schritt klären.")
    dots = dot_classes(row)
    return f"""<article class="match-card{' top' if index == 0 else ''}"><div class="match-head"><div><div class="rank">{'#1 · Top-Match' if index == 0 else '#' + str(index + 1)}</div><h3 class="pair">{h(pair)}</h3><div class="chips"><span class="chip">{score}/100</span><span class="chip stage">{h(stage_label(row, index))}</span></div></div><div class="fit-dots"><div class="fit-dot"><span class="dot {dots[0]}"></span>Problemfit</div><div class="fit-dot"><span class="dot {dots[1]}"></span>Anlass</div><div class="fit-dot"><span class="dot {dots[2]}"></span>Umsetzungsnähe</div></div></div><p class="summary">{h(summary)}</p><p class="next"><strong>Next Step:</strong> {h(next_line)}</p></article>"""


def icon(kind: str) -> str:
    if kind == "users":
        return '<svg viewBox="0 0 24 24"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><circle cx="9" cy="7" r="4" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>'
    if kind == "clock":
        return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M12 7v5l3 2" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>'
    return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8" fill="none" stroke="currentColor" stroke-width="1.8"/><circle cx="12" cy="12" r="3" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M12 2v3m0 14v3M2 12h3m14 0h3" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>'


def match_page(prefix: str) -> str:
    cards = "".join(match_card(row, i) for i, row in enumerate(read_matches()))
    current = load_current_history()
    created = clean_text(current.get("created_at")) or ""
    return f"""<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>SCIP Match Selector · SOR</title><link rel="icon" type="image/png" href="{prefix}sor-tab-o.png"><link rel="apple-touch-icon" href="{prefix}sor-tab-o.png"><style>{MATCH_CSS}</style></head><body><div class="shell">{header("scip", prefix)}<section class="meta-strip"><div class="meta-item"><span class="meta-icon">{icon("users")}</span>6 kuratierte Opportunity Matches</div><div class="meta-dot">·</div><div class="meta-item"><span class="meta-icon">{icon("clock")}</span>aktualisiert {h(created)}</div><div class="meta-dot">·</div><div class="meta-item"><span class="meta-icon">{icon("target")}</span>Detected actor bridges, evidence status and next validation steps.</div></section><main class="panel match-panel"><div class="title-row"><div class="title-left"><span class="target-icon">{icon("target")}</span><h1 class="section-title">SCIP Match Selector</h1></div><a class="button" href="zirp_meeting_recommendations_ai_latest.html">Alle Pairings bewerten</a></div><section class="match-grid">{cards}</section></main>{footer("Ready for feedback")}</div></body></html>"""


def redirect_page(target: str) -> str:
    return f'<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="refresh" content="0; url={h(target)}"><title>Weiterleitung · SOR</title><link rel="icon" type="image/png" href="sor-tab-o.png"><script>location.replace("{h(target)}");</script></head><body><p><a href="{h(target)}">Zur aktuellen Wochenanalyse</a></p></body></html>'


def crop_logo_assets() -> None:
    try:
        from PIL import Image
    except Exception:
        return
    for path in [REPORTS / "sor-logo-full.png", REPORTS / "sor-logo.png"]:
        if not path.exists():
            continue
        img = Image.open(path).convert("RGBA")
        px = img.load()
        w, ht = img.size
        coords = []
        for y in range(ht):
            for x in range(w):
                r, g, b, a = px[x, y]
                if a > 0 and (r > 72 or g > 72 or b > 92):
                    coords.append((x, y))
        if not coords:
            continue
        xs, ys = zip(*coords)
        pad_x = int(w * .035)
        pad_y = int(ht * .055)
        box = (max(min(xs) - pad_x, 0), max(min(ys) - pad_y, 0), min(max(xs) + pad_x, w), min(max(ys) + pad_y, ht))
        img.crop(box).save(path)


def write(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8", newline="\n")


def postprocess_site() -> None:
    crop_logo_assets()
    write(PROJECT_DIR / "index.html", home_page("zirp_berichte/"))
    write(PROJECT_DIR / "zirp_dashboard.html", home_page("zirp_berichte/"))
    write(REPORTS / "zirp_dashboard.html", home_page(""))
    write(REPORTS / "signals.html", redirect_page("zirp_dashboard.html"))
    write(REPORTS / "scip_archive.html", match_page(""))


if __name__ == "__main__":
    postprocess_site()
