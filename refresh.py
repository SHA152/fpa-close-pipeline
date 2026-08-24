#!/usr/bin/env python3
"""
==============================================================================
 fpa-close-pipeline / refresh.py
 AI-assisted monthly reporting refresh — May 2026 close (synthetic data)
 Author: Shakil Ahmad, CFA
==============================================================================

 DESIGN PRINCIPLE (mirrors the workflow diagram in README.md, stage for stage):

   The LLM NARRATES. It NEVER does arithmetic.
   Every number is computed deterministically in Stages 1-2, handed to the
   model as a sealed JSON payload in Stage 3, and every figure the model
   writes is checked back against that payload in Stage 5. A human approves
   the output through a pull-request review (the gate lives OUTSIDE this
   script, in GitHub - see README).

 STAGES (same names/numbers as the diagram):
   Stage 1  EXTRACT   - read the recalculated Excel model (values only)
   Stage 2  COMPUTE   - variances & aggregates in Python; cross-check vs model
   Stage 3  GROUND    - build the approved-figures payload; wire the prompt
   Stage 4  NARRATE   - one prompt template, run twice ({{audience}} variable)
   Stage 5  VALIDATE  - guardrail: reject any figure not in the payload
   Stage 6  PUBLISH   - render docs/index.html + append the run log
   Human gate         - git branch -> pull request -> review -> merge -> Pages

 USAGE:
   python refresh.py                    # full run, guardrail ON
   python refresh.py --guardrail off    # naive mode: model computes derived
                                        # figures itself (used once, to show
                                        # the review catching a wrong number)
   python refresh.py --dry-run          # no API call (plumbing test)

 The Anthropic API key is read from the ANTHROPIC_API_KEY environment
 variable (never hardcoded, never committed - see .env.example).
==============================================================================
"""

import argparse
import datetime
import json
import os
import re
import sys
from pathlib import Path

import openpyxl

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
ROOT = Path(__file__).parent
MODEL_XLSX = ROOT / "model" / "Blockstream_FPA_Model_Shakil_Ahmad.xlsx"
PROMPT_FILE = ROOT / "prompts" / "commentary_prompt.txt"
RUNS_DIR = ROOT / "runs"
DOCS_DIR = ROOT / "docs"

LLM_MODEL = "claude-sonnet-4-5"   # cost-efficient; swap to claude-haiku-4-5 for cheaper
MAX_TOKENS = 1000
TEMPERATURE = 0.2                  # low: commentary should be stable, not creative

AUDIENCES = ["bu_leader", "cfo_board"]

# Materiality threshold (mirrors Inputs!C37:C38 in the model)
MATERIAL_USD = 50      # $000s
MATERIAL_PCT = 0.10


# ============================================================================
# STAGE 1 - EXTRACT: read the recalculated model. Values only, no retyping.
# ============================================================================
def stage1_extract():
    """Pull raw actual/budget lines and the model's own computed anchors.

    We read TWO kinds of cells:
      (a) raw data lines from Data_Actuals / Data_Budget  -> inputs to Stage 2
      (b) the model's computed cells (BvA, Forecast, GL_Recon) -> used in
          Stage 2 as a CROSS-CHECK, so the script provably ties to the
          workbook a reviewer opens.
    """
    wb = openpyxl.load_workbook(MODEL_XLSX, data_only=True)
    A, B = wb["Data_Actuals"], wb["Data_Budget"]
    K = wb["Data_KPIs"]

    lines = {}          # line_item -> dict of monthly actual/budget
    for r in range(5, 24):                      # 19 P&L lines
        name = A.cell(r, 3).value
        lines[name] = {
            "segment": A.cell(r, 1).value,
            "category": A.cell(r, 2).value,
            "actual": [A.cell(r, c).value for c in range(4, 9)],   # Jan..May
            "budget": [B.cell(r, c).value for c in range(4, 9)],
        }

    kpis = {}
    for r in range(5, 20):
        kpis[K.cell(r, 1).value] = [K.cell(r, c).value for c in range(3, 8)]

    model_anchors = {                            # the workbook's own answers
        "may_total_revenue":  wb["BvA"]["B20"].value,
        "may_operating_profit": wb["BvA"]["B35"].value,
        "ytd_operating_profit": wb["BvA"]["G35"].value,
        "ytd_treasury_mark":  wb["BvA"]["G36"].value,
        "base_fy_operating_result": wb["Forecast"]["C5"].value,
        "bear_fy_operating_result": wb["Forecast"]["C6"].value,
        "swing": wb["Forecast"]["C7"].value,
        "btc_scenario_price": wb["Inputs"]["C5"].value,
        "corrected_may_operating_profit": wb["GL_Recon"]["B38"].value,
    }
    return lines, kpis, model_anchors


# ============================================================================
# STAGE 2 - COMPUTE: all arithmetic happens HERE, in deterministic Python.
# ============================================================================
def stage2_compute(lines, model_anchors):
    """Compute May & YTD variances, flag materiality, and cross-check that
    this script's math ties to the Excel model to the dollar. If the two
    disagree, the run ABORTS - numbers never reach the LLM unverified."""

    def var(a, b):
        d = a - b
        p = None if b == 0 else d / abs(b)
        return d, p

    rows = []
    rev_may = rev_ytd = cogs_may = cogs_ytd = opex_may = opex_ytd = 0.0
    for name, d in lines.items():
        may_a, may_b = d["actual"][4], d["budget"][4]
        ytd_a, ytd_b = sum(d["actual"]), sum(d["budget"])
        dv, dp = var(may_a, may_b)
        yv, yp = var(ytd_a, ytd_b)
        is_cost = d["category"] in ("COGS", "OpEx")
        material = abs(dv) >= MATERIAL_USD and (dp is not None and abs(dp) >= MATERIAL_PCT)
        rows.append({
            "line": name, "segment": d["segment"], "category": d["category"],
            "may_actual": may_a, "may_budget": may_b, "may_var": dv,
            "may_var_pct": None if dp is None else round(dp, 4),
            "ytd_actual": ytd_a, "ytd_budget": ytd_b, "ytd_var": yv,
            "material": material,
            "direction": ("favorable" if (dv < 0 if is_cost else dv > 0) else "unfavorable"),
        })
        if d["category"] == "Revenue":
            rev_may += may_a; rev_ytd += ytd_a
        elif d["category"] == "COGS":
            cogs_may += may_a; cogs_ytd += ytd_a
        elif d["category"] == "OpEx":
            opex_may += may_a; opex_ytd += ytd_a

    # budget-side aggregates (so commentary can quote budget and variance
    # totals without deriving anything)
    rev_may_b = sum(d["budget"][4] for d in lines.values() if d["category"] == "Revenue")
    rev_ytd_b = sum(sum(d["budget"]) for d in lines.values() if d["category"] == "Revenue")
    cogs_may_b = sum(d["budget"][4] for d in lines.values() if d["category"] == "COGS")
    cogs_ytd_b = sum(sum(d["budget"]) for d in lines.values() if d["category"] == "COGS")
    opex_may_b = sum(d["budget"][4] for d in lines.values() if d["category"] == "OpEx")
    opex_ytd_b = sum(sum(d["budget"]) for d in lines.values() if d["category"] == "OpEx")

    # segment aggregates - given to the model so it never needs to add lines
    def seg_sum(seg, cat, idx=None):
        vals = [d for d in lines.values() if d["segment"] == seg and d["category"] == cat]
        return (sum(d["actual"][idx] for d in vals) if idx is not None
                else sum(sum(d["actual"]) for d in vals))
    ah_rev_may, ah_cogs_may = seg_sum("App & Hardware", "Revenue", 4), seg_sum("App & Hardware", "COGS", 4)
    en_rev_may, en_cogs_may = seg_sum("Enterprise SaaS", "Revenue", 4), seg_sum("Enterprise SaaS", "COGS", 4)
    seg = {
        "app_hardware_revenue_may": ah_rev_may,
        "app_hardware_contribution_margin_may": ah_rev_may - ah_cogs_may,
        "enterprise_revenue_may": en_rev_may,
        "enterprise_contribution_margin_may": en_rev_may - en_cogs_may,
        "app_hardware_revenue_ytd": seg_sum("App & Hardware", "Revenue"),
        "enterprise_revenue_ytd": seg_sum("Enterprise SaaS", "Revenue"),
    }

    mtm = lines["BTC Treasury Mark-to-Market"]
    op_may, op_may_b = rev_may - cogs_may - opex_may, rev_may_b - cogs_may_b - opex_may_b
    op_ytd, op_ytd_b = rev_ytd - cogs_ytd - opex_ytd, rev_ytd_b - cogs_ytd_b - opex_ytd_b
    agg = {
        "segment_aggregates_may": seg,
        "may_total_revenue": rev_may, "may_total_revenue_budget": rev_may_b,
        "may_revenue_var": rev_may - rev_may_b,
        "may_gross_profit": rev_may - cogs_may,
        "may_total_opex": opex_may, "may_total_opex_budget": opex_may_b,
        "may_opex_var": opex_may - opex_may_b,
        "may_operating_profit": op_may, "may_operating_profit_budget": op_may_b,
        "may_operating_var": op_may - op_may_b,
        "may_treasury_mark": mtm["actual"][4],
        "ytd_total_revenue": rev_ytd, "ytd_total_revenue_budget": rev_ytd_b,
        "ytd_revenue_var": rev_ytd - rev_ytd_b,
        "ytd_operating_profit": op_ytd, "ytd_operating_profit_budget": op_ytd_b,
        "ytd_operating_var": op_ytd - op_ytd_b,
        "ytd_treasury_mark": sum(mtm["actual"]),
        "may_net_result_incl_treasury": op_may + mtm["actual"][4],
        "ytd_net_result_incl_treasury": op_ytd + sum(mtm["actual"]),
    }

    # ---- CROSS-CHECK: script math must tie to the workbook, to the dollar --
    for key in ("may_total_revenue", "may_operating_profit",
                "ytd_operating_profit", "ytd_treasury_mark"):
        if abs(agg[key] - model_anchors[key]) > 0.5:
            sys.exit(f"[ABORT] Cross-check failed on {key}: "
                     f"script={agg[key]} model={model_anchors[key]}. "
                     f"Numbers do not reach the LLM unverified.")
    print("[OK] Stage 2 cross-check: script arithmetic ties to the Excel "
          "model on all four anchor figures.")
    return rows, agg


# ============================================================================
# STAGE 3 - GROUND: seal the numbers into a payload; wire the prompt.
# ============================================================================
def stage3_ground(rows, agg, kpis, model_anchors, guardrail_on):
    """Build the ONLY numbers the model is allowed to use, and render the
    prompt template. The template has two variables:
        {{audience}}      - bu_leader | cfo_board (the brief's requirement)
        {{payload_json}}  - the injection mechanism: real May figures reach
                            the prompt at runtime as JSON, never retyped.
    Guardrail OFF (the naive run): derived figures (variances, YTD, totals)
    are withheld and the model is asked to work them out itself - exactly
    what a lazy implementation does, and what Stage 5 + human review catch."""

    payload = {
        "period": "May 2026 close (YTD = Jan-May 2026); all figures $000s",
        "materiality_rule": f"flag only if |$var| >= {MATERIAL_USD} AND |%var| >= {MATERIAL_PCT:.0%}",
        "kpis_jan_to_may": kpis,
        "gl_reconciliation": {
            "cloud_duplicate_invoice": {"je": "JE-1147", "amount": 30,
                "corrected_may_cloud_opex": 450,
                "note": "duplicate MongoDB invoice INV-MDB-771; ledger-side fix, reported P&L unaffected"},
            "travel_cutoff": {"je": "JE-1203", "amount": 150,
                "corrected_may_travel_events": 25,
                "note": "June conference sponsorship expensed in May; belongs in June as prepaid release"},
            "amp_rev_rec": {"je": "JE-1305", "amount": 132,
                "corrected_may_amp_revenue": 1248,
                "note": "annual license (term starts Jun) fully recognized in May; defer ratably at 11/mo"},
            "corrected_may_operating_profit": model_anchors["corrected_may_operating_profit"],
        },
        "scenario": {
            "btc_price_assumption": model_anchors["btc_scenario_price"],
            "base_fy_operating_result": round(model_anchors["base_fy_operating_result"], 1),
            "bear_fy_operating_result": round(model_anchors["bear_fy_operating_result"], 1),
            "swing": round(model_anchors["swing"], 1),
        },
    }
    # totals of the close adjustments, computed here so the model never sums
    _gl = payload["gl_reconciliation"]
    _gl["total_gross_adjustments_identified"] = (
        _gl["cloud_duplicate_invoice"]["amount"]
        + _gl["travel_cutoff"]["amount"] + _gl["amp_rev_rec"]["amount"])
    _gl["net_impact_on_may_operating_profit"] = (
        _gl["travel_cutoff"]["amount"] - _gl["amp_rev_rec"]["amount"])

    if guardrail_on:
        payload["aggregates"] = {k: (v if isinstance(v, dict) else round(v, 1))
                                 for k, v in agg.items()}
        payload["pnl_lines"] = rows                      # full precomputed set
    else:
        # NAIVE MODE: raw monthly values only - the model must derive
        # totals, YTD and variances itself. This is the run we expect
        # the validator and the human review to catch.
        payload["pnl_lines_raw_monthly"] = [
            {"line": r["line"], "may_actual": r["may_actual"],
             "may_budget": r["may_budget"]} for r in rows]
        payload["instruction_note"] = ("Derived figures were not provided; "
                                       "compute totals, YTD and variances yourself.")

    template = PROMPT_FILE.read_text(encoding="utf-8")
    return payload, template


# ============================================================================
# STAGE 4 - NARRATE: one template, two audiences, full evidence capture.
# ============================================================================
def stage4_narrate(template, payload, audience, run_id, dry_run):
    """Call the Anthropic API. The complete request AND response are written
    to runs/<run_id>_<audience>.json - that file is the execution evidence
    (timestamp, model, the injected May figures, and the raw output)."""
    prompt = (template
              .replace("{{audience}}", audience)
              .replace("{{payload_json}}", json.dumps(payload, indent=2)))

    if dry_run:
        text = (f"[DRY RUN - no API call] Narrative for {audience} would be "
                f"generated here from the injected payload.")
        usage = {}
    else:
        import anthropic
        client = anthropic.Anthropic()          # key from ANTHROPIC_API_KEY
        resp = client.messages.create(
            model=LLM_MODEL, max_tokens=MAX_TOKENS, temperature=TEMPERATURE,
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.content[0].text
        usage = {"input_tokens": resp.usage.input_tokens,
                 "output_tokens": resp.usage.output_tokens}

    evidence = {
        "run_id": run_id,
        "timestamp_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "model": LLM_MODEL,
        "audience": audience,
        "request": {"model": LLM_MODEL, "max_tokens": MAX_TOKENS,
                    "temperature": TEMPERATURE, "prompt": prompt},
        "response_text": text,
        "usage": usage,
    }
    out = RUNS_DIR / f"{run_id}_{audience}.json"
    out.write_text(json.dumps(evidence, indent=2), encoding="utf-8")
    print(f"[OK] Stage 4 ({audience}): response saved -> {out.name}")
    return text


# ============================================================================
# STAGE 5 - VALIDATE: the guardrail. Every figure the model wrote must exist
#                     in the approved payload. Anything else is a violation.
# ============================================================================
FIG_RE = re.compile(r"\(?\$?-?[\d,]+(?:\.\d+)?(?:\s?[KkMm](?![A-Za-z]))?\)?")
YEAR_RE = re.compile(r"^(19|20)\d{2}$")          # bare years are not figures
IDENT_RE = re.compile(r"[A-Za-z]-?$")            # JE-1305, INV-MDB-771 etc.

def _approved_values(payload):
    """Flatten every number in the payload into a set of allowed values."""
    vals = set()
    def walk(x):
        if isinstance(x, dict):
            for v in x.values(): walk(v)
        elif isinstance(x, list):
            for v in x: walk(v)
        elif isinstance(x, (int, float)) and x is not None:
            vals.add(round(float(x), 4))
    walk(payload)
    return vals

def _parse_number(tok):
    neg = "(" in tok or tok.strip().startswith("-")
    t = tok.replace("$", "").replace(",", "").replace("(", "").replace(")", "").strip()
    mult = 1.0
    if t and t[-1] in "KkMm":
        mult = 1000.0 if t[-1] in "Mm" else 1.0   # payload is already $000s
        t = t[:-1]
    try:
        v = float(t) * mult
        return -v if neg and v > 0 else v
    except ValueError:
        return None

def stage5_validate(text, payload, audience):
    """Return the list of figures in the narrative that do NOT tie to any
    approved payload value (tolerance +/-0.6 for rounding, +/-0.15pp for
    percentages). An empty list = the narrative is fully grounded."""
    approved = _approved_values(payload)
    approved_pct = {round(v * 100, 1) for v in approved if -5 <= v <= 5}
    violations = []
    for m in FIG_RE.finditer(text):
        tok = m.group(0)
        v = _parse_number(tok)
        if v is None or abs(v) < 3:          # ignore small counts ("3 issues")
            continue
        raw = tok.replace("(", "").replace(")", "").strip()
        if YEAR_RE.match(raw):               # "May 2026", "FY2026" - not figures
            continue
        if IDENT_RE.search(text[max(0, m.start()-4):m.start()]):
            continue                         # part of an identifier (JE-1305)
        is_pct = text[m.end():m.end()+1] == "%"
        bare = ("$" not in tok) and (tok.strip()[-1] not in "KkMm") and not is_pct
        if bare and abs(v) < 50:
            continue                         # bare small ints are counts, not figures
        # sign-tolerant ("a loss of $1,656K" for -1656) and unit-tolerant
        # ("$32.1M" for an ARR stored as 32.1 in $M). The scaled (/1000)
        # candidates use a MUCH tighter tolerance and only apply at >= 1,
        # so real errors can never hide behind the unit conversion.
        ok = (any(abs(c - a) <= 0.6 for a in approved for c in (v, -v))
              or any(abs(c) >= 1 and abs(c - a) <= max(0.05, 0.002 * abs(a))
                     for a in approved for c in (v / 1000.0, -v / 1000.0))
              or (is_pct and (any(abs(c - p) <= 0.5 for p in approved_pct for c in (v, -v))
                              or any(abs(c/100 - a) <= 0.002 for a in approved for c in (v, -v)))))
        if not ok:
            near = min(approved, key=lambda a: abs(v - a)) if approved else None
            violations.append({"figure_in_output": tok.strip(),
                               "parsed_value": v,
                               "nearest_approved_value": near})
    status = "PASS" if not violations else "FAIL"
    print(f"[{status}] Stage 5 validation ({audience}): "
          f"{len(violations)} ungrounded figure(s).")
    for v in violations:
        print(f"        -> '{v['figure_in_output']}' not in approved payload "
              f"(nearest approved: {v['nearest_approved_value']})")
    return violations


# ============================================================================
# STAGE 6 - PUBLISH: render the static dashboard + append the run log.
# ============================================================================
def stage6_publish(rows, agg, model_anchors, narratives, run_id, run_record):
    log_path = RUNS_DIR / "run_log.jsonl"
    with log_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(run_record) + "\n")
    html = render_dashboard(rows, agg, model_anchors, narratives, log_path)
    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")
    print(f"[OK] Stage 6: dashboard rendered -> docs/index.html ; "
          f"run appended -> runs/run_log.jsonl")


def _fmt(v):
    if v is None: return "n/m"
    return f"({abs(v):,.0f})" if v < 0 else f"{v:,.0f}"

def render_dashboard(rows, agg, anchors, narratives, log_path):
    tiles = [
        ("May Revenue", agg["may_total_revenue"], ""),
        ("May Operating Result", agg["may_operating_profit"], ""),
        ("YTD Operating Result", agg["ytd_operating_profit"], ""),
        ("YTD BTC Treasury Mark", agg["ytd_treasury_mark"], "below operating profit"),
        ("Base FY Operating Result", anchors["base_fy_operating_result"], f"BTC ${anchors['btc_scenario_price']:,.0f}"),
        ("Bear FY / Swing", anchors["bear_fy_operating_result"],
         f"swing {_fmt(anchors['swing'])}"),
    ]
    tile_html = "".join(
        f'<div class="tile"><div class="tl">{t}</div>'
        f'<div class="tv">{_fmt(v)}</div><div class="ts">{s}</div></div>'
        for t, v, s in tiles)

    trs = ""
    for r in rows:
        badge = ""
        if r["material"]:
            cls = "fav" if r["direction"] == "favorable" else "unfav"
            badge = f'<span class="badge {cls}">{r["direction"].title()} - material</span>'
        pct = "n/m" if r["may_var_pct"] is None else f"{r['may_var_pct']*100:,.1f}%"
        trs += (f'<tr><td>{r["line"]}</td><td>{r["segment"]}</td>'
                f'<td class="n">{_fmt(r["may_actual"])}</td>'
                f'<td class="n">{_fmt(r["may_budget"])}</td>'
                f'<td class="n">{_fmt(r["may_var"])}</td>'
                f'<td class="n">{pct}</td><td>{badge}</td></tr>')

    nar_html = ""
    for aud, label in (("bu_leader", "Business-Unit Leader view"),
                       ("cfo_board", "CFO / Board view")):
        body = narratives.get(aud, "(not generated this run)").replace("\n", "<br>")
        nar_html += f'<div class="panel"><h3>{label}</h3><div class="nar">{body}</div></div>'

    log_rows = ""
    if log_path.exists():
        for line in log_path.read_text(encoding="utf-8").strip().splitlines():
            e = json.loads(line)
            st = e.get("validation", "?")
            cls = "fav" if st == "PASS" else "unfav"
            log_rows += (f'<tr><td>{e.get("run_id","")}</td><td>{e.get("timestamp_utc","")[:19]}</td>'
                         f'<td>{e.get("model","")}</td><td>{e.get("guardrail","")}</td>'
                         f'<td><span class="badge {cls}">{st}</span></td></tr>')

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>May 2026 Close - Segment Performance</title>
<style>
  :root {{ --bg:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e; --line:#e3e2dd;
          --good:#006000; --goodbg:#e6f4ea; --bad:#b00000; --badbg:#fce4e4; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--ink);
         font:14px/1.5 Arial, "Helvetica Neue", sans-serif; padding:28px; }}
  .wrap {{ max-width:1080px; margin:0 auto; }}
  h1 {{ font-size:21px; margin:0 0 2px; }}
  .sub {{ color:var(--ink2); font-size:12.5px; margin-bottom:20px; }}
  .tiles {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
            gap:10px; margin-bottom:22px; }}
  .tile {{ border:1px solid var(--line); border-radius:8px; padding:12px 14px; background:#fff; }}
  .tl {{ font-size:11.5px; color:var(--ink2); }}
  .tv {{ font-size:20px; font-weight:bold; margin:2px 0; }}
  .ts {{ font-size:11px; color:var(--ink2); }}
  h2 {{ font-size:15px; margin:26px 0 8px; }}
  table {{ border-collapse:collapse; width:100%; background:#fff;
           border:1px solid var(--line); font-size:12.5px; }}
  th {{ text-align:left; background:#24292f; color:#fff; padding:6px 9px; font-weight:bold; }}
  td {{ padding:5px 9px; border-top:1px solid var(--line); }}
  td.n {{ text-align:right; font-variant-numeric:tabular-nums; }}
  .badge {{ font-size:11px; padding:1px 7px; border-radius:9px; white-space:nowrap; }}
  .badge.fav {{ background:var(--goodbg); color:var(--good); }}
  .badge.unfav {{ background:var(--badbg); color:var(--bad); }}
  .panels {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }}
  .panel {{ border:1px solid var(--line); border-radius:8px; background:#fff; padding:14px 16px; }}
  .panel h3 {{ margin:0 0 8px; font-size:13.5px; }}
  .nar {{ font-size:12.5px; color:#222; }}
  .foot {{ margin-top:26px; color:var(--ink2); font-size:11.5px;
           border-top:1px solid var(--line); padding-top:10px; }}
  @media (max-width:760px) {{ .panels {{ grid-template-columns:1fr; }} }}
  .scroll {{ overflow-x:auto; }}
</style></head><body><div class="wrap">
<h1>May 2026 Close &mdash; Segment Performance &amp; Bitcoin-Native KPIs</h1>
<div class="sub">Fully synthetic case-study data ($000s). Numbers computed
deterministically from the Excel model; AI narrates only; every figure
validated against the approved payload; published via reviewed pull request.</div>
<div class="tiles">{tile_html}</div>
<h2>May Actual vs Budget &mdash; materiality: &ge;$50K and &ge;10%</h2>
<div class="scroll"><table><tr><th>Line</th><th>Segment</th><th>Actual</th>
<th>Budget</th><th>$ Var</th><th>% Var</th><th>Flag</th></tr>{trs}</table></div>
<h2>AI Commentary &mdash; one prompt, two audiences</h2>
<div class="panels">{nar_html}</div>
<h2>Run log</h2>
<div class="scroll"><table><tr><th>Run</th><th>Timestamp (UTC)</th><th>Model</th>
<th>Guardrail</th><th>Validation</th></tr>{log_rows}</table></div>
<div class="foot">Pipeline: refresh.py (Stages 1-6) &middot; human gate: pull-request
review &amp; merge before publication &middot; BTC treasury mark reported below
operating profit (budgeted at zero by design) &middot; Shakil Ahmad, CFA</div>
</div></body></html>"""


# ============================================================================
# main - the stages in diagram order
# ============================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--guardrail", choices=["on", "off"], default="on")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    guardrail_on = args.guardrail == "on"

    run_id = datetime.datetime.now().strftime("run_%Y%m%d_%H%M%S")
    print(f"=== {run_id} | guardrail={args.guardrail} | model={LLM_MODEL} ===")

    lines, kpis, anchors = stage1_extract()                       # Stage 1
    rows, agg = stage2_compute(lines, anchors)                    # Stage 2
    payload, template = stage3_ground(rows, agg, kpis, anchors,   # Stage 3
                                      guardrail_on)
    narratives, all_violations = {}, {}
    for audience in AUDIENCES:                                    # Stage 4
        text = stage4_narrate(template, payload, audience, run_id, args.dry_run)
        narratives[audience] = text
        if not args.dry_run:
            all_violations[audience] = stage5_validate(text, payload, audience)  # Stage 5

    validation = ("DRY" if args.dry_run else
                  "PASS" if not any(all_violations.values()) else "FAIL")
    run_record = {"run_id": run_id,
                  "timestamp_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                  "model": LLM_MODEL, "guardrail": args.guardrail,
                  "validation": validation,
                  "violations": all_violations}
    # A FAILED validation still publishes the run LOG (the audit trail) but
    # refuses to publish the narratives to the dashboard.
    if validation == "FAIL":
        narratives = {a: "(withheld - failed figure validation; see run log)"
                      for a in narratives}
    stage6_publish(rows, agg, anchors, narratives, run_id, run_record)  # Stage 6

    print(f"=== done: validation={validation}. Next: git checkout -b {run_id}, "
          f"commit, open PR, review, merge (the human gate). ===")
    sys.exit(0 if validation != "FAIL" else 2)


if __name__ == "__main__":
    main()
