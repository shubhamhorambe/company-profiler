# =========================
# pipeline_core_v2.py
# PART 1 / 3 — Imports, Config, Helpers, Excel Writers
# =========================

import io
import os
import time
import json
import re
import math
from copy import copy
from datetime import date
from typing import Optional, Dict, Any, Tuple, List
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import requests
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openai import OpenAI

from mergermarket_import import MergermarketImport, parse_mergermarket_export
from mergermarket_browser import (
    MergermarketCredentials,
    MergermarketSearchPlan,
    fetch_mergermarket_transactions,
)
from research_agents import ResearchOrchestrator


# ---------------------------
# TEMPLATE / CONFIG
# ---------------------------
TEMPLATE_FILE = "Company_Input_v2.xlsx"  # must live next to this file
TEMPLATE_FILENAME = TEMPLATE_FILE  # backward-compat alias
SHEET_NAME = "Company and Industry"
COMPANY_NAME_CELL = "B1"
MA_SHEET = "M&A Landscape"
BUYERS_SHEET = "Prospective Buyers"
COMPETITORS_SHEET = "Competitors"
FINANCIALS_SHEET = "Financials"
CREDIT_RATING_SHEET = "Credit Rating"
SOURCE_DOSSIER_SHEET = "Source Dossier"
MM_RAW_SHEET = "MM Raw Export"
FINAL_COMPS_SHEET = "Final Comps"
ADJACENT_COMPS_SHEET = "Adjacent"
EXCLUDED_COMPS_SHEET = "Excluded"
COMPS_SOURCES_SHEET = "Sources"
COMPS_SEARCH_LOG_SHEET = "Search Log"
COMPS_QA_SHEET = "QA"
EVIDENCE_REGISTER_SHEET = "Evidence Register"

START_ROW = 5

# Main-sheet (NEW layout)
MAIN_COL_SECTION = 3   # C (merged for bucketed sections)
MAIN_COL_BUCKET = 4    # D
MAIN_COL_ANSWER = 5    # E
MAIN_COL_SOURCES = 6   # F


# Table sheets
COMP_COL_NAME = 1
COMP_COL_TYPE = 2
COMP_COL_REGION = 3
COMP_COL_HQ = 4
COMP_COL_DESC = 5
COMP_COL_SOURCES = 6

PRIVATECIRCLE_BASE_URL = "https://privatecircle.co"

RESEARCH_MODEL = os.getenv("OPENAI_RESEARCH_MODEL", "gpt-5.6-terra")
REWRITE_MODEL = os.getenv("OPENAI_REWRITE_MODEL", "gpt-5.6-luna")

MAX_DESC_CHARS = 2200
MAX_SOURCES_CHARS = 1500
SLEEP_BETWEEN_CALLS_SEC = 0.4


# ---------------------------
# BUCKET SPECS (what goes into column B)
# ---------------------------
BUCKET_ROW_SPECS: Dict[str, List[str]] = {
    "what does it exactly do and what are the products": [
        "Offering buckets",
        "Where it sits in customer workflow/value chain",
        "Monetisation model",
        "Key capabilities",
    ],
    "industries catered and customers": [
        "End industries (ranked)",
        "Customer segments / buyer persona",
        "Named customers (ONLY if disclosed)",
        "Geography (demand regions)",
    ],
    "industry landscape": [
        "Market size & growth (triangulated)",
        "Top players (Global)",
        "Top players (India)",
        "Key growth drivers",
        "Key entry barriers",
    ],
}


# ---------------------------
# Row prompts (kept close to your original intent)
# ---------------------------
ROW_PROMPTS: Dict[str, str] = {
    "about the company": """
Write an IM-ready company overview using publicly available information.

Cover:
- What the company does (1–2 lines)
- Key offerings and end-markets (high level)
- Headquarters / operating geography if available
- Scale indicators ONLY if explicitly stated
- 1–2 factual differentiators if visible

Rules:
- Use official websites, registries, databases, or press articles
- If information is partial, summarise what is known and clearly state
  “Not publicly disclosed” for gaps
- Do NOT leave the section empty unless absolutely no information exists
- No marketing language or assumptions

Output: 8–12 concise bullets
""".strip(),

    "what does it exactly do and what are the products": """
Explain what the company does and its offerings.

Structure:
- Offering buckets (3–6 buckets): bucket name + what it includes + typical use cases
- Where it sits in customer workflow/value chain
- Upstream/downstream integration points (generic)
- Monetisation model (if disclosed)
- Key capabilities (5–8, grouped; avoid SKU lists)

Rules:
- No repeating founding/HQ
- No exhaustive catalog dumps
""".strip(),

    "industries catered and customers": """
Identify end-markets served and customer profile.

Structure:
- End industries (ranked): Top 3–6 industries, 1 line each
- Customer segments / buyer persona
- Named customers (ONLY if disclosed): official site/press/regulatory/case studies
- Geography (demand regions) if disclosed

Rules:
- Do NOT infer customer names
""".strip(),

    "other relevant info": """
High-signal extras for investors (do not repeat earlier rows).

Structure:
- Certifications & compliance (ONLY if sourced)
- Technology / IP (ONLY if sourced)
- Delivery / service model (implementation, support, SLAs/AMCs; if disclosed)
- Key risks / constraints (ONLY if sourced)

Rules:
- Avoid generic claims
""".strip(),

    "industry landscape": """
Industry landscape (triangulated market size & structured sections).

Structure:
- Market size & growth (TRIANGULATE): range + base year + CAGR; use >=2 sources
- Top players (Global): top 5 (name + HQ)
- Top players (India): top 5 (name + HQ)
- Key growth drivers (4–6)
- Key entry barriers (4–6)

Rules:
- Neutral tone
""".strip(),

"competitors": """
Return ONLY valid JSON. No commentary.

OBJECTIVE:
Build the most exhaustive competitor universe for the target company, covering:
- Direct competitors (same product/service + same buyer budget)
- Indirect/substitute competitors (alternative ways to solve the same job-to-be-done)
- Adjacent peers (partial overlap, sometimes shortlisted together)
- Vertical/value-chain overlaps (upstream/downstream players that increasingly bundle/compete)

OUTPUT JSON:
{
  "direct_competitors": [
    {
      "name": "string",
      "hq": "City, Country/State or Not publicly disclosed",
      "competitor_type": "Direct",
      "overlap_dimension": "Product/Service | Customer segment | Geography | Channel | Value-chain position",
      "description": "1–2 lines: what they sell + exact overlap vs target",
      "sources": ["https://..."]
    }
  ],
  "adjacent_peers": [
    {
      "name": "string",
      "hq": "City, Country/State or Not publicly disclosed",
      "competitor_type": "Indirect/Substitute | Adjacent | Vertical overlap",
      "overlap_dimension": "Product/Service | Customer segment | Geography | Channel | Value-chain position",
      "description": "1–2 lines: why they compete/overlap (partial is okay)",
      "sources": ["https://..."]
    }
  ]
}

STRICT RULES:
- Maximise COVERAGE (India/local + global), but NO guessing.
- Do NOT include the target company itself.
- Do NOT include customers as competitors.
- Large conglomerates allowed ONLY if the competing business/division is identifiable.
- Each entry MUST have >=1 credible http(s) source URL.
- Do NOT invent URLs; use only URLs supported by evidence provided to you.
- If you cannot find a credible source, EXCLUDE that entry.
- If unsure whether direct or adjacent, put in adjacent_peers.
TARGET COUNTS (if available): Direct 8–20, Adjacent 8–25.
""".strip(),

"m&a landscape": """
Return ONLY valid JSON. No commentary.

OBJECTIVE:
Create an evidence-driven M&A landscape for the target's industry and adjacent industries, including:
- Horizontal consolidation deals (competitor buys competitor)
- Capability acquisitions (tech/product/engineering/design)
- Geographic expansion deals
- Vertical integration deals (upstream/downstream)
- PE platform creation + bolt-on rollups
- Strategic minority investments / JVs if common in the sector

OUTPUT JSON:
{
  "deals": [
    {
      "acquirer": "string",
      "target": "string",
      "year": "YYYY or Not disclosed",
      "deal_type": "Majority | Minority | Platform | Bolt-on | Strategic investment | JV",
      "deal_value": "string or Value not disclosed",
      "rationale": "1–2 lines: why this deal happened + relevance to the market",
      "source": "https://..."
    }
  ],
  "most_acquisitive_buyers": [
    {
      "name": "string",
      "pattern": "1 line: roll-up/platform/bolt-on strategy with evidence",
      "source": "https://..."
    }
  ],
  "activity_trend": {
    "summary": "2–4 lines: deal activity trend (deal-count/value proxy) and what it implies",
    "sources": ["https://...", "https://..."]
  }
}

STRICT RULES:
- Maximise COVERAGE (India/local + global), preferably last 7–10 years if available.
- Only include deals that clearly relate to the same or adjacent market.
- Do NOT include deals where acquirer or target is the target company itself unless explicitly evidenced as a real transaction.
- Every deal/buyer MUST have >=1 credible http(s) source.
- Do NOT invent URLs; use only URLs supported by evidence provided to you.
- If value/year not disclosed, keep the deal but explicitly mark as Not disclosed / Value not disclosed (still with source).
- Aim 15–40 deals if available; else return as many as evidenced.
""".strip(),


"prospective buyers": """
Return ONLY valid JSON. No commentary.

OBJECTIVE:
Identify the broadest realistic buyer universe for the target company, across:
- Strategic buyers (same market)
- Competitors as buyers (horizontal consolidation)
- Adjacent strategics (capability/segment adjacency)
- Vertical integrators (upstream/downstream integration)
- Financial buyers (PE/Infra/Family office) and PE-backed platforms with acquisition appetite
- Global + domestic buyers

A buyer is "realistic" if there is EVIDENCE of at least one:
(1) Prior acquisitions in the same/adjacent space,
(2) Ownership/portfolio in the space,
(3) Stated inorganic growth intent,
(4) Clear strategic adjacency (capability/geography/customer access) supported by sources.

OUTPUT JSON:
{
  "buyers": [
    {
      "buyer": "string",
      "category": "Strategic | Competitor | Adjacent Strategic | Vertical Integrator | PE/Infra Platform | Family Office",
      "hq": "City, Country/State or Not disclosed",
      "rationale": "1–2 lines: specific buyer logic (market/capability/geography/vertical integration/roll-up)",
      "frictions": "1 line: key risk (antitrust/overlap/integration/regulatory) or Not material",
      "prior_acquisitions": "1 line: 0–2 precedent examples if available; else Not disclosed",
      "source": "https://..."
    }
  ]
}

STRICT RULES:
- Maximise COVERAGE (India/local + global), but NO guessing.
- Do NOT include the target company itself.
- Do NOT include entities only because they are large/famous; must have evidence.
- Each buyer MUST have >=1 credible http(s) source URL.
- Do NOT invent URLs; use only URLs supported by evidence provided to you.
- If you cannot find a credible source, EXCLUDE that buyer.
- Minimum 15 buyers if the market supports it; aim 25–60 if available.
""".strip(),

}

DEFAULT_ROW_PROMPT = """
Provide a detailed, IM-ready analysis relevant to the given topic.

Rules:
- Be factual and evidence-driven
- Use publicly available sources
- Do NOT invent data
- Clearly state “Not publicly disclosed” where applicable
""".strip()


# ---------------------------
# SAFE TEMPLATE LOADER
# ---------------------------
def load_template_workbook() -> openpyxl.Workbook:
    here = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(here, TEMPLATE_FILE)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found at: {template_path}")
    return openpyxl.load_workbook(template_path)


# ---------------------------
# Helpers
# ---------------------------
def clamp(s: str, n: int) -> str:
    s = (s or "").strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def strip_markdown_links(text: str) -> str:
    return re.sub(r"\[([^\]]+)\]\((https?://[^)]+)\)", r"\1 (\2)", text or "")


def canonical_url(u: str) -> str:
    try:
        u = (u or "").strip()
        parts = urlsplit(u)
        q = [
            (k, v)
            for k, v in parse_qsl(parts.query, keep_blank_values=True)
            if k.lower()
            not in {"utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "gclid", "fbclid"}
        ]
        return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(q, doseq=True), ""))
    except Exception:
        return (u or "").strip()


def extract_urls(text: str) -> List[str]:
    urls = re.findall(r"https?://[^\s)>\]]+", text or "")
    seen, out = set(), []
    for u in urls:
        u = canonical_url(u.strip().rstrip(".,;"))
        if u and u not in seen:
            seen.add(u)
            out.append(u)
    return out


def extract_json_object(text: str) -> Dict[str, Any]:
    """
    Robust JSON extraction: tries full parse; else extracts first balanced {...} object.
    """
    text = (text or "").strip()
    if not text:
        return {}

    try:
        obj = json.loads(text)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        pass

    start = text.find("{")
    if start == -1:
        return {}

    depth = 0
    in_str = False
    esc = False

    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    block = text[start : i + 1]
                    try:
                        obj = json.loads(block)
                        return obj if isinstance(obj, dict) else {}
                    except Exception:
                        return {}
    return {}


def normalize_bullets(desc: Any) -> str:
    if isinstance(desc, list):
        desc = "\n".join([str(x).strip() for x in desc if str(x).strip()])

    desc = strip_markdown_links(str(desc or "").strip())

    # handle python-list-ish string
    if "['" in desc and "']" in desc:
        parts = re.findall(r"'([^']+)'", desc)
        if parts:
            desc = "\n".join([f"- {p.strip()}" for p in parts if p.strip()])

    lines = [ln.strip() for ln in desc.splitlines() if ln.strip()]
    if lines and not any(lines[0].startswith(ch) for ch in ("-", "•", "*")):
        lines = [f"- {ln}" for ln in lines]

    seen = set()
    cleaned = []
    for ln in lines:
        key = re.sub(r"\s+", " ", ln).strip().lower()
        if key not in seen:
            seen.add(key)
            cleaned.append(ln)

    return "\n".join(cleaned)


def get_row_prompt(particular: str) -> str:
    key = (particular or "").strip().lower()
    return ROW_PROMPTS.get(key, DEFAULT_ROW_PROMPT)


# ---------------------------
# EXCEL STYLE: SAFE ROW STYLE COPY (NO _style)
# ---------------------------
def copy_row_style_safe(ws, src_row: int, dst_row: int, max_col: int = 5):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)


# ---------------------------
# NEW: Bucketed matrix writer (A merged, B labels, C answers, E sources)
# ---------------------------
def _estimate_row_height_for_text(text: str, col_width_chars: float, base: float = 14.0, max_h: float = 240.0) -> float:
    """Rough auto-fit for wrapped Excel cells (openpyxl cannot truly auto-fit row height).
    We approximate number of wrapped lines from column width and newline breaks.
    """
    if not text:
        return base
    s = str(text)
    # count explicit lines first
    hard_lines = s.split("\n")
    lines = 0
    # guard for very small widths
    w = max(8.0, float(col_width_chars or 30.0))
    for hl in hard_lines:
        # ~1 char ~= 1 "width unit" in Excel column width for typical fonts
        l = max(1, math.ceil(len(hl) / w))
        lines += l
    return min(max_h, max(base, base * lines))


def write_bucketed_matrix(
    ws,
    start_row: int,
    section_title: str,
    rows: List[Dict[str, Any]],
    sources: List[str],
    col_title: int = MAIN_COL_SECTION,   # C (Category)
    col_label: int = MAIN_COL_BUCKET,    # D (Sub-categories)
    col_value: int = MAIN_COL_ANSWER,    # E (Description)
    col_sources: int = MAIN_COL_SOURCES, # F (Source)
) -> int:
    """
    IN-PLACE writer for Company_Input_v2.xlsx.

    IMPORTANT:
    - Does NOT insert rows
    - Does NOT merge/unmerge cells
    - Preserves all existing formatting
    - Ensures the template's sub-category labels (column D) are present for EVERY row,
      including the first row of each bucket (e.g., 'Offering buckets').
    - Writes sources per sub-category row (same source list repeated; upstream prompts currently return a shared list)
    """
    if not isinstance(rows, list):
        rows = []
    if not isinstance(sources, list):
        sources = extract_urls(str(sources))

    # Ensure we have the expected label order for this bucket
    key = (section_title or "").strip().lower()
    required_labels = BUCKET_ROW_SPECS.get(key, [])
    if required_labels:
        # Build a map from label -> bullets if provided
        by_label: Dict[str, List[str]] = {}
        for rr in rows:
            lab = str(rr.get("label", "")).strip()
            bullets = rr.get("bullets", [])
            if not isinstance(bullets, list):
                bullets = [str(bullets)]
            bullets = [str(x).strip() for x in bullets if str(x).strip()]
            if lab:
                by_label[lab] = bullets

        normalized_rows: List[Dict[str, Any]] = []
        for lab in required_labels:
            bullets = by_label.get(lab, [])
            if not bullets:
                bullets = ["Not publicly disclosed"]
            normalized_rows.append({"label": lab, "bullets": bullets})
        rows = normalized_rows
    else:
        # fallback
        for rr in rows:
            rr["label"] = str(rr.get("label", "")).strip()

    n = len(rows)

    # Prepare sources string (repeat per sub-row)
    src_urls = [canonical_url(u) for u in sources if isinstance(u, str) and u.startswith("http")]
    src_urls = list(dict.fromkeys(src_urls))[:8]
    src_str = "\n".join(src_urls)

    # Column width (for row height estimation). If None, Excel default ~8.43.
    col_letter = openpyxl.utils.get_column_letter(col_value)
    col_w = ws.column_dimensions[col_letter].width or 30.0

    for i, rr in enumerate(rows):
        r = start_row + i

        # 1) Keep category in column C only on the first row (template already does this)
        if i == 0:
            ws.cell(r, col_title).value = section_title

        # 2) Ensure sub-category label is written (fixes missing first-row titles)
        lab = str(rr.get("label", "")).strip()
        if lab:
            ws.cell(r, col_label).value = lab

        # 3) Write description (bullets)
        bullets = rr.get("bullets", [])
        if not isinstance(bullets, list):
            bullets = [str(bullets)]
        bullets = [str(x).strip() for x in bullets if str(x).strip()]
        ws.cell(r, col_value).value = normalize_bullets("\n".join(bullets))

        # 4) Write sources per sub-category row
        if src_str:
            ws.cell(r, col_sources).value = src_str

        # 5) Ensure readable formatting (wrap + top align)
        for c in (col_label, col_value, col_sources):
            cell = ws.cell(r, c)
            # preserve existing styles where possible; only enforce wrap/top
            align = cell.alignment.copy(wrapText=True, vertical="center")
            cell.alignment = align

        # 6) Approximate auto-fit row height for Description
        val = ws.cell(r, col_value).value
        ws.row_dimensions[r].height = _estimate_row_height_for_text(val, col_w)

    return n
def _style_table(ws, header_row: int, max_col: int):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_row = ws.max_row

    for c in range(1, max_col + 1):
        cell = ws.cell(header_row, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = border

    for r in range(header_row + 1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.alignment = body_align
            cell.border = border

    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{max_row}"


def format_competitors_sheet(ws):
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 55
    ws.column_dimensions["F"].width = 50
    _style_table(ws, header_row=1, max_col=6)


def format_ma_sheet(ws):
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 60
    ws.column_dimensions["F"].width = 50
    _style_table(ws, header_row=1, max_col=6)

    for r in range(2, ws.max_row + 1):
        ws.cell(r, 3).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def format_buyers_sheet(ws):
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 60
    ws.column_dimensions["E"].width = 38
    ws.column_dimensions["F"].width = 42
    ws.column_dimensions["G"].width = 50
    _style_table(ws, header_row=1, max_col=7)


def format_credit_rating_sheet(ws):
    ws.freeze_panes = "A2"
    widths = [18, 15, 18, 22, 18, 28, 18, 55, 48, 48, 55]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    _style_table(ws, header_row=1, max_col=len(widths))


def format_source_dossier_sheet(ws):
    ws.freeze_panes = "A2"
    widths = [22, 24, 48, 16, 65, 55]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    _style_table(ws, header_row=1, max_col=len(widths))


# ---------------------------
# ENSURE / CLEAR SHEETS
# ---------------------------
def ensure_competitors_sheet(wb: openpyxl.Workbook):
    ws = wb[COMPETITORS_SHEET] if COMPETITORS_SHEET in wb.sheetnames else wb.create_sheet(COMPETITORS_SHEET)
    if ws.max_row < 1 or ws.cell(1, 1).value is None:
        headers = ["Competitor", "Type", "Region", "HQ", "Description", "Sources"]
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c).value = h
    return ws


def ensure_ma_sheet(wb: openpyxl.Workbook):
    ws = wb[MA_SHEET] if MA_SHEET in wb.sheetnames else wb.create_sheet(MA_SHEET)
    if ws.max_row < 1 or ws.cell(1, 1).value is None:
        headers = ["Acquirer", "Target", "Year", "Deal Value", "Rationale", "Source"]
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c).value = h
    return ws


def ensure_buyers_sheet(wb: openpyxl.Workbook):
    ws = wb[BUYERS_SHEET] if BUYERS_SHEET in wb.sheetnames else wb.create_sheet(BUYERS_SHEET)
    if ws.max_row < 1 or ws.cell(1, 1).value is None:
        headers = ["Buyer", "Category", "HQ", "Rationale", "Frictions", "Prior Acquisitions", "Source"]
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c).value = h
    return ws


def ensure_financials_sheet(wb: openpyxl.Workbook):
    ws = wb[FINANCIALS_SHEET] if FINANCIALS_SHEET in wb.sheetnames else wb.create_sheet(FINANCIALS_SHEET)
    if ws.max_row > 1:
        ws.delete_rows(1, ws.max_row)
    return ws


def ensure_credit_rating_sheet(wb: openpyxl.Workbook):
    ws = wb[CREDIT_RATING_SHEET] if CREDIT_RATING_SHEET in wb.sheetnames else wb.create_sheet(CREDIT_RATING_SHEET)
    headers = [
        "Agency",
        "Report Date",
        "Rating Action",
        "Rating",
        "Outlook",
        "Instrument",
        "Amount",
        "Rationale",
        "Key Strengths",
        "Key Risks",
        "Official Report URL",
    ]
    if ws.max_row < 1 or ws.cell(1, 1).value is None:
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c).value = h
    return ws


def ensure_source_dossier_sheet(wb: openpyxl.Workbook):
    ws = wb[SOURCE_DOSSIER_SHEET] if SOURCE_DOSSIER_SHEET in wb.sheetnames else wb.create_sheet(SOURCE_DOSSIER_SHEET)
    headers = ["Source Type", "Publisher", "Title", "Published Date", "URL", "Notes"]
    if ws.max_row < 1 or ws.cell(1, 1).value is None:
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c).value = h
    return ws


def write_source_dossier(ws, dossier: Dict[str, Any]):
    row = ws.max_row + 1
    official_website = dossier.get("official_website")
    if official_website:
        ws.cell(row, 1).value = "official_website"
        ws.cell(row, 2).value = dossier.get("legal_name") or "Company"
        ws.cell(row, 3).value = "Official company website"
        ws.cell(row, 4).value = "Not disclosed"
        ws.cell(row, 5).value = official_website
        ws.cell(row, 6).value = dossier.get("identity_notes") or "Identity-matched official website"
        row += 1

    seen = {official_website} if official_website else set()
    for source in dossier.get("sources", []):
        if not isinstance(source, dict) or source.get("url") in seen:
            continue
        seen.add(source.get("url"))
        ws.cell(row, 1).value = source.get("source_type")
        ws.cell(row, 2).value = source.get("publisher")
        ws.cell(row, 3).value = source.get("title")
        ws.cell(row, 4).value = source.get("published_date")
        ws.cell(row, 5).value = source.get("url")
        ws.cell(row, 6).value = source.get("notes")
        row += 1


def write_credit_reports(ws, reports: List[Dict[str, str]]):
    row = ws.max_row + 1
    fields = [
        "agency",
        "report_date",
        "rating_action",
        "rating",
        "outlook",
        "instrument",
        "amount",
        "rationale",
        "key_strengths",
        "key_risks",
        "report_url",
    ]
    for report in reports:
        if not isinstance(report, dict):
            continue
        for column, field in enumerate(fields, start=1):
            ws.cell(row, column).value = report.get(field)
        row += 1


COMPS_HEADERS = [
    "Announced Date",
    "Target",
    "Acquirer",
    "Deal Value (USD mn)",
    "Deal %",
    "Seller",
    "EV/EBITDA",
    "Mergermarket Deal ID",
    "Relevance Tier",
    "Relevance Rationale",
    "Transaction Type",
    "Status",
    "Original Deal Value",
    "Currency",
    "Enterprise Value (USD mn)",
    "Revenue (USD mn)",
    "EBITDA (USD mn)",
    "EV/Revenue",
    "Multiple Status",
    "Mergermarket Profile",
    "Description",
    "Target Sector",
    "Target Geography",
]


def _fresh_sheet(wb: openpyxl.Workbook, name: str):
    if name in wb.sheetnames:
        index = wb.sheetnames.index(name)
        wb.remove(wb[name])
        return wb.create_sheet(name, index)
    return wb.create_sheet(name)


def _safe_excel_value(value: Any):
    if isinstance(value, str):
        return f"'{value}" if value.startswith(("=", "+", "@")) else value
    if isinstance(value, (int, float, bool, time.struct_time)) or value is None:
        return value
    if hasattr(value, "isoformat"):
        return value
    return str(value)


def _format_comps_table(ws):
    header_fill = PatternFill("solid", fgColor="17365D")
    header_font = Font(bold=True, color="FFFFFF")
    header_border = Border(bottom=Side(style="medium", color="17365D"))
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = header_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "D2"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 1)}"
    widths = {
        "A": 14, "B": 28, "C": 28, "D": 18, "E": 12, "F": 24, "G": 14,
        "H": 20, "I": 14, "J": 48, "K": 20, "L": 16, "M": 20, "N": 11,
        "O": 20, "P": 18, "Q": 18, "R": 14, "S": 17, "T": 48, "U": 70,
        "V": 28, "W": 22,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 1).number_format = "yyyy-mm-dd"
        for column in (4, 13, 15, 16, 17):
            ws.cell(row, column).number_format = '$#,##0.0;[Red]($#,##0.0);-'
            ws.cell(row, column).alignment = Alignment(horizontal="right", vertical="top")
        ws.cell(row, 5).number_format = "0.0%;[Red](0.0%);-"
        for column in (7, 18):
            ws.cell(row, column).number_format = "0.0x;[Red](0.0x);-"
        for column in range(1, ws.max_column + 1):
            if column not in {4, 5, 7, 13, 15, 16, 17, 18}:
                ws.cell(row, column).alignment = Alignment(vertical="top", wrap_text=True)
        if ws.cell(row, 20).value:
            ws.cell(row, 20).font = Font(color="FF0000", underline="single")
        tier = str(ws.cell(row, 9).value or "")
        tier_fill = {
            "Core": "C6EFCE",
            "Broader": "DDEBF7",
            "Adjacent": "FFF2CC",
            "Excluded": "F4CCCC",
            "Unscreened": "FCE4D6",
        }.get(tier)
        if tier_fill:
            ws.cell(row, 9).fill = PatternFill("solid", fgColor=tier_fill)
            ws.cell(row, 9).font = Font(bold=True)
        long_text = max(len(str(ws.cell(row, 10).value or "")), len(str(ws.cell(row, 21).value or "")))
        ws.row_dimensions[row].height = min(90, max(18, 15 * math.ceil(long_text / 70)))


def _append_comp_row(ws, deal: Dict[str, Any]):
    row = ws.max_row + 1
    explicit_ev_ebitda = deal.get("ev_ebitda_reported")
    explicit_ev_revenue = deal.get("ev_revenue_reported")
    values = [
        deal.get("announced_date"),
        deal.get("target"),
        deal.get("acquirer") or "Not disclosed",
        deal.get("deal_value_usd_mn"),
        deal.get("deal_percent"),
        deal.get("seller") or "Not disclosed",
        explicit_ev_ebitda,
        deal.get("deal_id") or "Not disclosed",
        deal.get("relevance_tier") or "Unscreened",
        deal.get("relevance_rationale") or "Screening was not completed",
        deal.get("deal_type") or "Not disclosed",
        deal.get("deal_status") or "Not disclosed",
        deal.get("deal_value_original") or deal.get("deal_value_original_text"),
        deal.get("currency") or "Not disclosed",
        deal.get("enterprise_value_usd_mn"),
        deal.get("revenue_usd_mn"),
        deal.get("ebitda_usd_mn"),
        explicit_ev_revenue,
        "Reported" if explicit_ev_ebitda is not None else "Not computable",
        deal.get("profile_url"),
        deal.get("description"),
        deal.get("target_sector"),
        deal.get("target_geography"),
    ]
    for column, value in enumerate(values, start=1):
        ws.cell(row, column).value = _safe_excel_value(value)
    if explicit_ev_ebitda is None and deal.get("enterprise_value_usd_mn") is not None and deal.get("ebitda_usd_mn") not in (None, 0):
        ws.cell(row, 7).value = f'=IFERROR(O{row}/Q{row},"")'
        ws.cell(row, 19).value = "Calculated"
    if explicit_ev_revenue is None and deal.get("enterprise_value_usd_mn") is not None and deal.get("revenue_usd_mn") not in (None, 0):
        ws.cell(row, 18).value = f'=IFERROR(O{row}/P{row},"")'


def write_transaction_comps_workbook(
    wb: openpyxl.Workbook,
    imported: MergermarketImport,
    company_name: str,
    search_scope: str,
    search_queries: Optional[List[str]] = None,
    retrieval_mode: str = "Authorised Mergermarket export",
):
    raw = _fresh_sheet(wb, MM_RAW_SHEET)
    for column, header in enumerate(imported.headers, start=1):
        raw.cell(1, column).value = header
    for row_index, raw_row in enumerate(imported.raw_rows, start=2):
        for column, value in enumerate(raw_row, start=1):
            raw.cell(row_index, column).value = _safe_excel_value(value)
    raw.freeze_panes = "A2"
    raw.sheet_view.showGridLines = False
    raw.auto_filter.ref = f"A1:{get_column_letter(max(raw.max_column, 1))}{max(raw.max_row, 1)}"
    for cell in raw[1]:
        cell.fill = PatternFill("solid", fgColor="595959")
        cell.font = Font(bold=True, color="FFFFFF")
    for column in range(1, raw.max_column + 1):
        raw.column_dimensions[get_column_letter(column)].width = 22

    final_ws = _fresh_sheet(wb, FINAL_COMPS_SHEET)
    adjacent_ws = _fresh_sheet(wb, ADJACENT_COMPS_SHEET)
    excluded_ws = _fresh_sheet(wb, EXCLUDED_COMPS_SHEET)
    for ws in (final_ws, adjacent_ws, excluded_ws):
        for column, header in enumerate(COMPS_HEADERS, start=1):
            ws.cell(1, column).value = header

    for deal in imported.deals:
        tier = deal.get("relevance_tier")
        if tier in {"Core", "Broader"}:
            destination = final_ws
        elif tier == "Adjacent":
            destination = adjacent_ws
        else:
            destination = excluded_ws
        _append_comp_row(destination, deal)

    for ws in (final_ws, adjacent_ws, excluded_ws):
        _format_comps_table(ws)

    sources = _fresh_sheet(wb, COMPS_SOURCES_SHEET)
    source_headers = ["Mergermarket Deal ID", "Target", "Field", "Source Class", "URL", "Access Notes"]
    for column, header in enumerate(source_headers, start=1):
        sources.cell(1, column).value = header
    source_row = 2
    for deal in imported.deals:
        if not deal.get("profile_url"):
            continue
        values = [deal.get("deal_id"), deal.get("target"), "Seed transaction record", "Mergermarket", deal.get("profile_url"), "Imported from authorised export"]
        for column, value in enumerate(values, start=1):
            sources.cell(source_row, column).value = value
        sources.cell(source_row, 5).font = Font(color="FF0000", underline="single")
        source_row += 1
    format_source_dossier_sheet(sources)

    search_log = _fresh_sheet(wb, COMPS_SEARCH_LOG_SHEET)
    search_headers = ["Date", "Subject Company", "Search / Source", "Scope", "Result", "Unresolved / Next Step"]
    for column, header in enumerate(search_headers, start=1):
        search_log.cell(1, column).value = header
    search_log.append([
        time.strftime("%Y-%m-%d"),
        company_name,
        f"{retrieval_mode}: {imported.source_sheet}",
        search_scope,
        f"{len(imported.deals)} unique candidate transactions imported",
        "Run public-source enrichment for missing values and financial denominators",
    ])
    for query in search_queries or []:
        search_log.append([
            time.strftime("%Y-%m-%d"),
            company_name,
            "Automatic Mergermarket Deals search",
            search_scope,
            query,
            "Public-web coverage pass retained for complementary transactions and source validation",
        ])
    for warning in imported.warnings:
        search_log.append([time.strftime("%Y-%m-%d"), company_name, "Import QA", search_scope, warning, "Review source export layout if material"])
    _style_table(search_log, header_row=1, max_col=6)
    for column, width in enumerate([14, 28, 48, 28, 45, 55], start=1):
        search_log.column_dimensions[get_column_letter(column)].width = width

    qa = _fresh_sheet(wb, COMPS_QA_SHEET)
    qa.append(["Check", "Result", "Status", "Notes"])
    qa.append(["Unique imported transactions", len(imported.deals), "OK" if imported.deals else "FAIL", "After Deal ID and composite-key deduplication"])
    qa.append(["Core + Broader transactions", max(final_ws.max_row - 1, 0), "OK" if final_ws.max_row > 1 else "REVIEW", "Final Comps population"])
    qa.append(["Adjacent transactions", max(adjacent_ws.max_row - 1, 0), "OK", "Context-only transactions"])
    qa.append(["Excluded / unscreened transactions", max(excluded_ws.max_row - 1, 0), "REVIEW" if excluded_ws.max_row > 1 else "OK", "Review all unscreened records"])
    missing_dates = sum(1 for deal in imported.deals if not deal.get("announced_date"))
    missing_values = sum(1 for deal in imported.deals if deal.get("deal_value_usd_mn") is None)
    qa.append(["Missing announced dates", missing_dates, "OK" if missing_dates == 0 else "REVIEW", "Research missing dates before final use"])
    qa.append(["Missing USD deal values", missing_values, "OK" if missing_values == 0 else "REVIEW", "Do not treat funding amount as total deal value"])
    qa.append(["Coverage saturation", "Not run", "OPEN", "Requires two public-source gap passes with no additional relevant deal"])
    _style_table(qa, header_row=1, max_col=4)
    qa.column_dimensions["A"].width = 34
    qa.column_dimensions["B"].width = 22
    qa.column_dimensions["C"].width = 14
    qa.column_dimensions["D"].width = 65
    for row in range(2, qa.max_row + 1):
        status = str(qa.cell(row, 3).value or "")
        fill = {
            "OK": "C6EFCE",
            "REVIEW": "FFF2CC",
            "OPEN": "FCE4D6",
            "FAIL": "F4CCCC",
        }.get(status)
        if fill:
            qa.cell(row, 3).fill = PatternFill("solid", fgColor=fill)
            qa.cell(row, 3).font = Font(bold=True)


def write_evidence_register(
    wb: openpyxl.Workbook,
    facts_cache: Dict[Tuple[str, str], List[Dict[str, Any]]],
):
    ws = _fresh_sheet(wb, EVIDENCE_REGISTER_SHEET)
    headers = ["Topic", "Claim", "Evidence Quality", "Source URL"]
    for column, header in enumerate(headers, start=1):
        ws.cell(1, column).value = header
    seen = set()
    row = 2
    quality_order = {"regulator": 0, "official": 1, "industry_report": 2, "news": 3, "database": 4, "other": 5}
    records = []
    for (_, topic), facts in facts_cache.items():
        for fact in facts:
            key = (topic, fact.get("claim"), fact.get("evidence_url"))
            if key in seen:
                continue
            seen.add(key)
            records.append((topic, fact))
    records.sort(key=lambda item: (item[0], quality_order.get(item[1].get("evidence_quality"), 9)))
    for topic, fact in records:
        ws.cell(row, 1).value = topic
        ws.cell(row, 2).value = fact.get("claim")
        ws.cell(row, 3).value = fact.get("evidence_quality") or "other"
        ws.cell(row, 4).value = fact.get("evidence_url")
        ws.cell(row, 4).font = Font(color="FF0000", underline="single")
        row += 1
    _style_table(ws, header_row=1, max_col=4)
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 90
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 65


def clear_table_sheet_keep_header(ws):
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)


# =========================
# pipeline_core_v2.py
# PART 2 / 3 — PrivateCircle + Research (facts→rewrite) + JSON tables
# =========================

# ---------------------------
# PrivateCircle API helpers
# ---------------------------

def privatecircle_search_dropdown(query: str, token: str, limit: int = 15) -> List[Dict[str, Any]]:
    q = (query or "").strip()
    if len(q) < 2:
        return []

    data = pc_get(
        "/api/v1/companies/search/",
        token,
        params={"generic_search": f"\"{q}\"", "page": 1},
    )
    results = data.get("results", []) or []

    if not results:
        data = pc_get(
            "/api/v1/companies/search/",
            token,
            params={"name": q, "page": 1},
        )
        results = data.get("results", []) or []

    out = []
    for r in results[:limit]:
        out.append({
            "id": r.get("id"),
            "name": r.get("name"),
            "dba_name": r.get("dba_name"),
            "cin": r.get("cin_number") or r.get("cin"),
            "city": r.get("city"),
            "state": r.get("state"),
        })

    return [x for x in out if x.get("id") and x.get("name")]


def _pc_headers(token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def pc_get(path: str, token: str, params: Optional[Dict[str, Any]] = None, retries: int = 2) -> Dict[str, Any]:
    url = f"{PRIVATECIRCLE_BASE_URL}{path}"
    last_err = None
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=_pc_headers(token), params=params, timeout=30)
        except requests.RequestException as exc:
            last_err = exc
            if attempt >= retries:
                break
            time.sleep(0.6 * (attempt + 1))
            continue

        if r.status_code == 429 or r.status_code >= 500:
            last_err = RuntimeError(f"PrivateCircle HTTP {r.status_code}: {r.text[:800]}")
            if attempt >= retries:
                break
            time.sleep(1.5 * (attempt + 1))
            continue
        if r.status_code >= 400:
            # Authentication, permission, and bad-request errors are not transient.
            raise RuntimeError(f"PrivateCircle HTTP {r.status_code}: {r.text[:800]}")
        try:
            payload = r.json()
        except ValueError as exc:
            last_err = RuntimeError(f"PrivateCircle returned invalid JSON: {exc}")
            if attempt >= retries:
                break
            time.sleep(0.6 * (attempt + 1))
            continue
        if not isinstance(payload, dict):
            raise RuntimeError("PrivateCircle returned an unexpected non-object response.")
        return payload
    raise RuntimeError(f"PrivateCircle request failed: {last_err}")


def pc_get_company_endpoint(encrypted_id: str, suffix: str, token: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    try:
        return pc_get(f"/api/v1/companies/{encrypted_id}/{suffix}/", token, params=params)
    except RuntimeError as e:
        msg = str(e)
        if "HTTP 404" in msg or "Not Found" in msg:
            return pc_get(f"/api/v1/company/{encrypted_id}/{suffix}/", token, params=params)
        raise


# ---------------------------
# Financials (PrivateCircle)
# ---------------------------
def _pc_number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().lower()
    if text in {"", "na", "n/a", "nm", "none", "null", "-"}:
        return None
    try:
        return float(text.replace(",", ""))
    except (TypeError, ValueError):
        return None


def to_inr_mn(value: Any, divisor: float = 1_000_000.0) -> Optional[float]:
    """Convert an API amount to INR mn using the response-specific divisor."""
    number = _pc_number(value)
    if number is None:
        return None
    safe_divisor = divisor if divisor and divisor > 0 else 1.0
    return round(number / safe_divisor, 2)


def _pc_notes(payload: Dict[str, Any]) -> List[str]:
    notes = payload.get("notes") or []
    if isinstance(notes, str):
        return [notes]
    return [str(note) for note in notes if note]


def privatecircle_amount_divisor(payload: Dict[str, Any]) -> float:
    """Return the divisor needed to express PrivateCircle amounts in INR mn.

    PrivateCircle's current contract says statement values are already in INR mn,
    while older examples in the same documentation show rupee-scale integers. The
    response notes (or the explicit environment override) therefore take priority.
    """
    override = os.getenv("PC_FINANCIAL_INPUT_UNIT", "auto").strip().lower()
    if override in {"inr_mn", "mn", "million", "millions"}:
        return 1.0
    if override in {"inr", "rupees", "raw_inr"}:
        return 1_000_000.0

    unit_candidates = [payload.get("unit"), payload.get("units"), payload.get("currency_unit")]
    meta = payload.get("meta") if isinstance(payload.get("meta"), dict) else {}
    unit_candidates.extend([meta.get("unit"), meta.get("units"), meta.get("currency_unit")])
    unit_text = " ".join(str(value) for value in unit_candidates if value).lower()
    notes_text = " ".join(_pc_notes(payload)).lower()
    combined = f"{unit_text} {notes_text}"
    if re.search(r"inr\s*(mn|mm|million)", combined):
        return 1.0
    if "rupee" in combined or re.search(r"\binr\b", combined):
        return 1_000_000.0

    # Last-resort compatibility for legacy responses with no unit metadata.
    amount_values: List[float] = []
    excluded_fragments = ("year", "percent", "ratio", "days", "date", "id", "eps")
    for item in payload.get("items") or []:
        if not isinstance(item, dict):
            continue
        for key, value in item.items():
            if any(fragment in str(key).lower() for fragment in excluded_fragments):
                continue
            number = _pc_number(value)
            if number is not None and number != 0:
                amount_values.append(abs(number))
    if amount_values:
        amount_values.sort()
        median = amount_values[len(amount_values) // 2]
        if median >= 1_000_000:
            return 1_000_000.0
    return 1.0


def _pc_overview_item(payload: Dict[str, Any]) -> Dict[str, Any]:
    items = payload.get("items") or []
    return items[0] if items and isinstance(items[0], dict) else {}


def _privatecircle_preferred_source(overview: Dict[str, Any]) -> str:
    listing_status = re.sub(r"[^a-z]", "", str(overview.get("listing_status") or "").lower())
    return "exchange_filings" if listing_status in {"listed", "publiclylisted"} else "mca"


def _statement_candidates(preferred_source: str, requested_basis: str) -> List[Tuple[Optional[str], str]]:
    bases = [requested_basis] if requested_basis in {"consolidated", "standalone"} else ["consolidated", "standalone"]
    candidates: List[Tuple[Optional[str], str]] = []
    for source in (preferred_source, None):
        for basis in bases:
            candidate = (source, basis)
            if candidate not in candidates:
                candidates.append(candidate)
    return candidates


def _pull_privatecircle_dataset(
    encrypted_id: str,
    token: str,
    suffix: str,
    no_of_years: int,
    candidates: List[Tuple[Optional[str], str]],
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    errors: List[str] = []
    for source, basis in candidates:
        params: Dict[str, Any] = {
            "no_of_years": no_of_years,
            "xbrl_type_tag": basis,
        }
        if source:
            params["source"] = source
        if suffix != "basic-financial-ratios":
            params["type"] = "detailed"
        try:
            payload = pc_get_company_endpoint(encrypted_id, suffix, token, params=params)
        except RuntimeError as exc:
            message = str(exc)
            if any(code in message for code in ("HTTP 400", "HTTP 404", "HTTP 422")):
                errors.append(message)
                continue
            raise
        if payload.get("items"):
            return payload, {
                "source": source or "api_default",
                "xbrl_type_tag": basis,
                "item_count": len(payload.get("items") or []),
                "errors": errors,
            }
    fallback_basis = candidates[0][1] if candidates else "auto"
    return {"items": [], "notes": []}, {
        "source": preferred_source,
        "xbrl_type_tag": fallback_basis,
        "item_count": 0,
        "errors": errors,
    }


def pull_financial_statements(
    encrypted_id: str,
    token: str,
    no_of_years: int = 5,
    xbrl_type_tag: str = "auto",
) -> Dict[str, Any]:
    """Pull a complete, internally documented PrivateCircle financial package."""
    overview_payload = pc_get_company_endpoint(encrypted_id, "overview", token)
    overview = _pc_overview_item(overview_payload)
    preferred_source = _privatecircle_preferred_source(overview)
    candidates = _statement_candidates(preferred_source, xbrl_type_tag.strip().lower())

    income, income_meta = _pull_privatecircle_dataset(
        encrypted_id, token, "income-statement", no_of_years, candidates
    )
    chosen = (income_meta.get("source"), income_meta.get("xbrl_type_tag"))
    chosen_source = None if chosen[0] == "api_default" else chosen[0]
    ordered_candidates = [(chosen_source, chosen[1])] + [c for c in candidates if c != (chosen_source, chosen[1])]

    balance, balance_meta = _pull_privatecircle_dataset(
        encrypted_id, token, "balance-sheet", no_of_years, ordered_candidates
    )
    cash_flow, cash_flow_meta = _pull_privatecircle_dataset(
        encrypted_id, token, "cash-flow", no_of_years, ordered_candidates
    )
    ratios, ratios_meta = _pull_privatecircle_dataset(
        encrypted_id, token, "basic-financial-ratios", no_of_years, ordered_candidates
    )

    return {
        "overview": overview_payload,
        "income_statement": income,
        "balance_sheet": balance,
        "cash_flow": cash_flow,
        "basic_financial_ratios": ratios,
        "_meta": {
            "preferred_source": preferred_source,
            "income_statement": income_meta,
            "balance_sheet": balance_meta,
            "cash_flow": cash_flow_meta,
            "basic_financial_ratios": ratios_meta,
        },
    }


def _year_map(items: List[Dict[str, Any]]) -> Dict[int, Dict[str, Any]]:
    mapped: Dict[int, Dict[str, Any]] = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        try:
            year = int(item.get("fiscal_year"))
        except (TypeError, ValueError):
            continue
        mapped[year] = item
    return mapped


def _first_number(row: Dict[str, Any], fields: Tuple[str, ...]) -> Optional[float]:
    for field in fields:
        value = _pc_number(row.get(field))
        if value is not None:
            return value
    return None


def apply_financials_format(ws, max_row: int, n_years: int, section_rows: List[int], header_rows: List[int], check_rows: List[int]):
    last_col = 1 + max(n_years, 1)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B9"
    ws.column_dimensions["A"].width = 42
    for column in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(column)].width = 15

    ws["A1"].font = Font(bold=True, size=15, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    for column in range(2, last_col + 1):
        ws.cell(1, column).fill = PatternFill("solid", fgColor="17365D")

    thin_gray = Side(style="thin", color="D9E2F3")
    for row in range(1, max_row + 1):
        ws.row_dimensions[row].height = {2: 24, 3: 24, 4: 24, 5: 38, 6: 38, 7: 24}.get(row, 19)
        for column in range(1, last_col + 1):
            cell = ws.cell(row, column)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(
                horizontal="left" if column == 1 else "right",
                vertical="center",
                wrap_text=column == 1,
            )
            if cell.value is not None and row not in section_rows:
                cell.border = Border(bottom=thin_gray)

    for row in section_rows:
        for column in range(1, last_col + 1):
            cell = ws.cell(row, column)
            cell.fill = PatternFill("solid", fgColor="1F4E79")
            cell.font = Font(bold=True, color="FFFFFF")
        ws.row_dimensions[row].height = 22

    for row in header_rows:
        for column in range(1, last_col + 1):
            cell = ws.cell(row, column)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.font = Font(bold=True, color="17365D")
            cell.border = Border(bottom=Side(style="medium", color="7F8FA6"))

    for row in check_rows:
        ws.cell(row, 1).font = Font(bold=True)
        for column in range(2, last_col + 1):
            ws.cell(row, column).fill = PatternFill("solid", fgColor="E2F0D9")

    for row in range(2, min(max_row, 7) + 1):
        ws.cell(row, 1).font = Font(bold=True, color="404040")
        ws.cell(row, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def write_financials_sheet(wb: openpyxl.Workbook, encrypted_id: str, token: str):
    ws = ensure_financials_sheet(wb)
    fin = pull_financial_statements(encrypted_id, token, no_of_years=5, xbrl_type_tag="auto")

    datasets = {
        "income_statement": fin.get("income_statement") or {},
        "balance_sheet": fin.get("balance_sheet") or {},
        "cash_flow": fin.get("cash_flow") or {},
        "basic_financial_ratios": fin.get("basic_financial_ratios") or {},
    }
    overview = _pc_overview_item(fin.get("overview") or {})
    maps = {name: _year_map(payload.get("items") or []) for name, payload in datasets.items()}
    all_years = sorted({year for mapped in maps.values() for year in mapped}, reverse=True)[:5]
    years = sorted(all_years)
    if not years:
        ws["A1"] = "No financial statement data available from PrivateCircle."
        ws["A2"] = "The company match succeeded, but the statement endpoints returned no fiscal-year records."
        return

    n_years = len(years)
    last_col = 1 + n_years
    metadata = fin.get("_meta") or {}
    divisors = {
        name: privatecircle_amount_divisor(payload)
        for name, payload in datasets.items()
        if name != "basic_financial_ratios"
    }
    unit_descriptions = []
    for name in ("income_statement", "balance_sheet", "cash_flow"):
        if datasets[name].get("items"):
            interpretation = "already INR mn" if divisors[name] == 1 else "raw INR divided by 1,000,000"
            unit_descriptions.append(f"{name.replace('_', ' ').title()}: {interpretation}")

    basis_parts = []
    for name in ("income_statement", "balance_sheet", "cash_flow", "basic_financial_ratios"):
        item_meta = metadata.get(name) or {}
        if item_meta.get("item_count"):
            basis_parts.append(
                f"{name.replace('_', ' ').title()}: {item_meta.get('source')} / {item_meta.get('xbrl_type_tag')}"
            )

    ws["A1"] = "PrivateCircle Financial Analysis"
    ws["A2"] = "Company"
    ws["B2"] = overview.get("name") or "Not disclosed"
    ws["A3"] = "CIN"
    ws["B3"] = overview.get("true_cin") or overview.get("true_cid") or "Not disclosed"
    ws["A4"] = "Listing status"
    ws["B4"] = overview.get("listing_status") or "Not disclosed"
    ws["A5"] = "Statement source / filing basis"
    ws["B5"] = "; ".join(basis_parts) or "No statement metadata returned"
    ws["A6"] = "Units shown"
    ws["B6"] = "INR mn. " + ("; ".join(unit_descriptions) or "No amount-bearing statements returned")
    ws["A7"] = "Profile"
    ws["B7"] = overview.get("profile_link") or "PrivateCircle API"
    if last_col > 2:
        for metadata_row in range(2, 8):
            ws.merge_cells(start_row=metadata_row, start_column=2, end_row=metadata_row, end_column=last_col)

    section_rows: List[int] = []
    header_rows: List[int] = []
    check_rows: List[int] = []
    row_map: Dict[str, int] = {}
    r = 9

    def section(title: str, value_header: bool = False):
        nonlocal r
        ws.cell(r, 1).value = title
        section_rows.append(r)
        r += 1
        ws.cell(r, 1).value = "Metric"
        if value_header:
            ws.cell(r, 2).value = "Value"
        else:
            for j, year in enumerate(years, start=2):
                ws.cell(r, j).value = f"FY{year}"
        header_rows.append(r)
        r += 1

    def write_amount(label: str, dataset_name: str, *fields: str, sum_fields: Tuple[str, ...] = ()):
        nonlocal r
        ws.cell(r, 1).value = label
        divisor = divisors.get(dataset_name, 1.0)
        for j, year in enumerate(years, start=2):
            source_row = maps[dataset_name].get(year, {})
            raw = _first_number(source_row, tuple(fields))
            if raw is None and sum_fields:
                components = [_pc_number(source_row.get(field)) for field in sum_fields]
                available = [value for value in components if value is not None]
                raw = sum(available) if available else None
            ws.cell(r, j).value = to_inr_mn(raw, divisor)
            ws.cell(r, j).number_format = '#,##0.0;[Red](#,##0.0);-'
        row_map[label] = r
        r += 1

    def write_number(label: str, dataset_name: str, *fields: str, number_format: str = "0.0"):
        nonlocal r
        ws.cell(r, 1).value = label
        for j, year in enumerate(years, start=2):
            ws.cell(r, j).value = _first_number(maps[dataset_name].get(year, {}), tuple(fields))
            ws.cell(r, j).number_format = number_format
        row_map[label] = r
        r += 1

    def write_percent(label: str, dataset_name: str, *fields: str):
        nonlocal r
        ws.cell(r, 1).value = label
        for j, year in enumerate(years, start=2):
            value = _first_number(maps[dataset_name].get(year, {}), tuple(fields))
            ws.cell(r, j).value = value / 100.0 if value is not None else None
            ws.cell(r, j).number_format = '0.0%;[Red](0.0%);-'
        row_map[label] = r
        r += 1

    def write_formula_ratio(label: str, numerator: str, denominator: str):
        nonlocal r
        ws.cell(r, 1).value = label
        for j in range(2, last_col + 1):
            col = get_column_letter(j)
            ws.cell(r, j).value = f'=IFERROR({col}{row_map[numerator]}/{col}{row_map[denominator]},"")'
            ws.cell(r, j).number_format = '0.0%;[Red](0.0%);-'
        row_map[label] = r
        r += 1

    section("Income Statement")
    write_amount("Revenue from Operations", "income_statement", "revenue_from_operations", "total_sales_wo_other_income")
    write_amount("Other Income", "income_statement", "other_income")
    write_amount(
        "Total Revenue",
        "income_statement",
        "total_revenue",
        "total_income",
        sum_fields=("revenue_from_operations", "other_income"),
    )
    write_amount(
        "Export Sales",
        "income_statement",
        "total_export_sales",
        sum_fields=("export_goods_manf_sales", "export_goods_traded_sales", "export_services_sales"),
    )
    write_formula_ratio("Exports % of Revenue", "Export Sales", "Revenue from Operations")
    write_amount("Cost of Goods Sold", "income_statement", "cost_of_goods_sold")
    write_amount("Gross Profit", "income_statement", "gross_profit")
    write_formula_ratio("Gross Margin", "Gross Profit", "Revenue from Operations")
    write_amount("Operating EBITDA", "income_statement", "operating_ebitda")
    write_formula_ratio("Operating EBITDA Margin", "Operating EBITDA", "Revenue from Operations")
    write_amount("Reported EBITDA", "income_statement", "ebitda")
    write_amount("Depreciation & Amortisation", "income_statement", "depreciation_amortization")
    write_amount("Operating EBIT", "income_statement", "operating_ebit")
    write_amount("Finance Cost", "income_statement", "finance_cost")
    write_amount("Profit Before Tax", "income_statement", "pbt", "pbt_from_co")
    write_amount("Profit After Tax", "income_statement", "pat", "profit_loss_from_co")
    write_formula_ratio("PAT Margin", "Profit After Tax", "Total Revenue")
    r += 1

    section("Balance Sheet")
    write_amount("Equity Capital", "balance_sheet", "bs_total_equity_capital")
    write_amount("Reserves & Surplus", "balance_sheet", "bs_reserves_and_surplus")
    write_amount("Non-current Liabilities", "balance_sheet", "bs_noncurrent_liabilities")
    write_amount("Net Fixed Assets", "balance_sheet", "bs_total_net_fixed_assets")
    write_amount("Trade Receivables", "balance_sheet", "bs_trade_receivables")
    write_amount("Cash & Bank Balance", "balance_sheet", "bs_cash_and_bank_balance")
    write_amount("Current Assets ex Cash", "balance_sheet", "bs_current_assets_ex_cash")
    write_amount("Current Liabilities", "balance_sheet", "bs_current_liabilities")
    write_amount("Trade Payables", "balance_sheet", "bs_trade_payables")
    write_amount("Net Working Capital ex Cash", "balance_sheet", "bs_net_working_capital")
    write_amount("Total Assets", "balance_sheet", "bs_total_assets", "bs_total_applications")
    write_amount("Total Liabilities", "balance_sheet", "bs_total_liabilities")
    write_number("Receivable Days", "balance_sheet", "sales_outstanding_days")
    write_number("Inventory Days", "balance_sheet", "inventory_outstanding_days")
    write_number("Payable Days", "balance_sheet", "payable_outstanding_days")
    write_number("Cash Conversion Cycle", "balance_sheet", "cash_conversion_cycle")
    r += 1

    section("Cash Flow Statement")
    write_amount("Cash Flow from Operations", "cash_flow", "net_cashflows_from_operatng_activts", "net_cashflows_from_operations")
    write_amount("Cash Flow from Investing", "cash_flow", "net_cashflows_from_investing_activities")
    write_amount("Cash Flow from Financing", "cash_flow", "net_cashflows_from_financing_activities")
    write_amount("Purchase of Tangible Assets", "cash_flow", "purchase_of_tangible_assets")
    write_amount(
        "Purchase of Intangible Assets",
        "cash_flow",
        sum_fields=("prchs_of_intangible_assets_clsfdas_inv_acts", "purchase_of_intangible_assets_under_development"),
    )
    write_amount("Income Taxes Paid", "cash_flow", "income_taxes_paid_refund")
    write_amount("Net Change in Cash", "cash_flow", "net_inc_dec_in_cash_and_cash_equival")
    write_amount("Closing Cash & Cash Equivalents", "cash_flow", "cash_and_cash_equival_stat_at_end_of_period", "cash_and_bank_balance")
    ws.cell(r, 1).value = "Free Cash Flow (CFO - Capex)"
    for j in range(2, last_col + 1):
        col = get_column_letter(j)
        ws.cell(r, j).value = (
            f'=IFERROR({col}{row_map["Cash Flow from Operations"]}'
            f'-ABS({col}{row_map["Purchase of Tangible Assets"]})'
            f'-ABS({col}{row_map["Purchase of Intangible Assets"]}),"")'
        )
        ws.cell(r, j).number_format = '#,##0.0;[Red](#,##0.0);-'
    row_map["Free Cash Flow (CFO - Capex)"] = r
    r += 2

    section("Basic Financial Ratios")
    write_number("Current Ratio", "basic_financial_ratios", "current_ratio", number_format='0.0x;[Red](0.0x);-')
    write_number("Debt / Equity", "basic_financial_ratios", "debt_equity_ratio", number_format='0.0x;[Red](0.0x);-')
    write_percent("ROCE", "basic_financial_ratios", "roce_percent")
    write_percent("ROE", "basic_financial_ratios", "roe_percent")
    write_percent("ROI", "basic_financial_ratios", "roi_percent")
    r += 1

    section("Key Outputs", value_header=True)
    latest_col = get_column_letter(last_col)
    ws.cell(r, 1).value = "Latest Revenue"
    ws.cell(r, 2).value = f'={latest_col}{row_map["Revenue from Operations"]}'
    ws.cell(r, 2).number_format = '#,##0.0;[Red](#,##0.0);-'
    r += 1
    ws.cell(r, 1).value = "Latest Operating EBITDA Margin"
    ws.cell(r, 2).value = f'={latest_col}{row_map["Operating EBITDA Margin"]}'
    ws.cell(r, 2).number_format = '0.0%;[Red](0.0%);-'
    r += 1
    ws.cell(r, 1).value = "Latest PAT Margin"
    ws.cell(r, 2).value = f'={latest_col}{row_map["PAT Margin"]}'
    ws.cell(r, 2).number_format = '0.0%;[Red](0.0%);-'
    r += 1
    if n_years >= 4:
        start_index = n_years - 4
        start_col_letter = get_column_letter(2 + start_index)
        periods = years[-1] - years[start_index]
        ws.cell(r, 1).value = f"Revenue CAGR ({periods} years)"
        ws.cell(r, 2).value = (
            f'=IFERROR(POWER({latest_col}{row_map["Revenue from Operations"]}/'
            f'{start_col_letter}{row_map["Revenue from Operations"]},1/{periods})-1,"")'
        )
        ws.cell(r, 2).number_format = '0.0%;[Red](0.0%);-'
        r += 1
    r += 1

    section("Data Quality Checks")
    ws.cell(r, 1).value = "Balance Sheet Difference"
    for j in range(2, last_col + 1):
        col = get_column_letter(j)
        ws.cell(r, j).value = f'=IFERROR({col}{row_map["Total Assets"]}-{col}{row_map["Total Liabilities"]},"")'
        ws.cell(r, j).number_format = '#,##0.0;[Red](#,##0.0);-'
    check_rows.append(r)
    r += 1
    ws.cell(r, 1).value = "Statement Coverage"
    for j, year in enumerate(years, start=2):
        present = [name for name, mapped in maps.items() if year in mapped]
        ws.cell(r, j).value = f"{len(present)}/4 datasets"
        ws.cell(r, j).alignment = Alignment(horizontal="center")
    check_rows.append(r)
    r += 2

    ws.cell(r, 1).value = "PrivateCircle API Notes"
    section_rows.append(r)
    r += 1
    for dataset_name, payload in datasets.items():
        for note in _pc_notes(payload):
            ws.cell(r, 1).value = dataset_name.replace("_", " ").title()
            ws.cell(r, 2).value = note
            ws.cell(r, 2).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.row_dimensions[r].height = 30
            r += 1

    apply_financials_format(ws, max_row=r, n_years=n_years, section_rows=section_rows, header_rows=header_rows, check_rows=check_rows)


# ---------------------------
# OpenAI response helper (supports older SDKs)
# ---------------------------
def _responses_create(client: OpenAI, **kwargs):
    """
    Normalize older ``response_format`` calls to the Responses API ``text``
    format and remove parameters unsupported by current reasoning models.
    """
    response_format = kwargs.pop("response_format", None)
    if response_format and "text" not in kwargs:
        kwargs["text"] = {"format": response_format}

    model = str(kwargs.get("model", ""))
    if model.startswith("gpt-5"):
        kwargs.pop("temperature", None)

    try:
        return client.responses.create(**kwargs)
    except TypeError:
        kwargs.pop("text", None)
        return client.responses.create(**kwargs)


def repair_json_once(bad_text: str, schema_hint: str, client: OpenAI) -> Dict[str, Any]:
    prompt = f"""
Fix the following into valid JSON that matches this schema exactly.

SCHEMA:
{schema_hint}

BAD_TEXT:
{bad_text}

Return ONLY valid JSON.
""".strip()
    rr = _responses_create(
        client,
        model=REWRITE_MODEL,
        input=prompt,
        max_output_tokens=1200,
        temperature=0,
    )
    return extract_json_object(rr.output_text or "")


# ---------------------------
# Market definition (web_search)
# ---------------------------
def build_market_definition_prompt(company: str) -> str:
    return f"""
Identify the industry that the company belongs to.

Company: {company}

Return ONLY JSON:
{{
  "primary_market": "industry / market category name",
  "market_description": "1–2 lines: what is included/excluded",
  "include_geographies": ["India","Global"],
  "keywords": ["8–14 search keywords for sizing + competitors"],
  "exclude_markets": ["umbrella markets to avoid (optional)"],
  "anchor_sources": ["2–5 URLs"]
}}

Rules:
- Use web search. No guessing.
- Prefer a buyer- and benchmarking-relevant industry label (not tech-micro-niche).
""".strip()


def get_market_definition(company: str, client: OpenAI) -> Dict[str, Any]:
    resp = _responses_create(
        client,
        model=RESEARCH_MODEL,
        input=build_market_definition_prompt(company),
        tools=[{"type": "web_search"}],
        response_format={"type": "json_object"},
        max_output_tokens=1400,
        temperature=0,
    )
    raw = (resp.output_text or "").strip()
    data = extract_json_object(raw)
    if not isinstance(data, dict) or not data.get("primary_market"):
        schema_hint = """{ "primary_market":"...", "market_description":"...", "include_geographies":["India","Global"], "keywords":["..."], "exclude_markets":["..."], "anchor_sources":["https://..."] }"""
        data = repair_json_once(raw, schema_hint, client)
    if not isinstance(data, dict) or not data.get("primary_market"):
        return {}
    if "exclude_markets" not in data or not isinstance(data["exclude_markets"], list):
        data["exclude_markets"] = []
    if "include_geographies" not in data or not isinstance(data["include_geographies"], list):
        data["include_geographies"] = ["India", "Global"]
    if "keywords" not in data or not isinstance(data["keywords"], list):
        data["keywords"] = []
    return data


# ---------------------------
# Facts prompt
# ---------------------------
def build_fact_prompt(company: str, particular: str, market_def: Dict[str, Any]) -> str:
    row_prompt = get_row_prompt(particular)
    key = (particular or "").strip().lower()

    market_context = ""
    if key in {"industry landscape", "competitors", "m&a landscape", "prospective buyers"} and market_def:
        market_context = f"""
MARKET DEFINITION (MANDATORY — do not deviate):
- Primary market: {market_def.get("primary_market")}
- Description: {market_def.get("market_description")}
- Geography scope: {", ".join(market_def.get("include_geographies", []))}
- Keywords: {", ".join(market_def.get("keywords", []))}
- EXCLUDE: {", ".join(market_def.get("exclude_markets", []))}
""".strip()

    identity_context = ""
    if market_def.get("official_website") or market_def.get("legal_name"):
        verified_sources = market_def.get("verified_sources") or []
        identity_context = f"""
ENTITY IDENTITY (use this to avoid similarly named companies):
- Legal name: {market_def.get("legal_name") or company}
- Official website: {market_def.get("official_website") or "Not confirmed"}
- Verified high-value sources: {", ".join(verified_sources[:12]) or "None pre-verified"}
""".strip()

    return f"""
Use web search and collect ONLY verifiable facts.

Company: {company}
Topic: {particular}

{market_context}
{identity_context}

Row-specific instructions:
{row_prompt}

SOURCE RULES:
- Prefer official sites, regulator filings, reputable business news, investor/IR pages, credible reports.
- Avoid low-quality directories and SEO blogs unless unavoidable.

Return ONLY JSON:
{{
  "facts": [
    {{
      "claim": "short factual statement (no fluff)",
      "evidence_url": "https://...",
      "evidence_quality": "official|industry_report|regulator|news|database|other"
    }}
  ]
}}

Rules:
- Every claim MUST have an evidence_url.
- No guessing / no inference.
""".strip()


def fetch_facts_cached(
    company: str,
    particular: str,
    market_def: Dict[str, Any],
    client: OpenAI,
    cache: Dict[Tuple[str, str], List[Dict[str, Any]]],
) -> List[Dict[str, Any]]:
    k = (company.strip().lower(), particular.strip().lower())
    if k in cache:
        return cache[k]

    fact_prompt = build_fact_prompt(company, particular, market_def)
    r1 = _responses_create(
        client,
        model=RESEARCH_MODEL,
        input=fact_prompt,
        tools=[{"type": "web_search"}],
        response_format={"type": "json_object"},
        max_output_tokens=2600,
        temperature=0,
    )

    raw = (r1.output_text or "").strip()
    data = extract_json_object(raw)
    if not data or "facts" not in data:
        schema_hint = """{ "facts":[{"claim":"...","evidence_url":"https://...","evidence_quality":"official|industry_report|regulator|news|database|other"}]}"""
        data = repair_json_once(raw, schema_hint, client)

    facts = data.get("facts", [])
    cleaned: List[Dict[str, Any]] = []
    if isinstance(facts, list):
        for f in facts:
            if not isinstance(f, dict):
                continue
            claim = str(f.get("claim", "")).strip()
            url = canonical_url(str(f.get("evidence_url", "")).strip())
            qual = str(f.get("evidence_quality", "")).strip().lower()
            if claim and url.startswith("http"):
                cleaned.append({"claim": claim, "evidence_url": url, "evidence_quality": qual})

    cache[k] = cleaned
    return cleaned


# ---------------------------
# Flat IM prompt (for About the company + other non-bucket rows)
# ---------------------------
def build_im_prompt_flat(particular: str, facts_json: str, market_def: Dict[str, Any]) -> str:
    row_prompt = get_row_prompt(particular)
    key = (particular or "").strip().lower()

    market_context = ""
    if key in {"industry landscape", "competitors", "m&a landscape", "prospective buyers"} and market_def:
        market_context = f"""
MARKET DEFINITION (MANDATORY — do not deviate):
Primary market: {market_def.get("primary_market")}
Description: {market_def.get("market_description")}
Geography: {", ".join(market_def.get("include_geographies", []))}
Keywords: {", ".join(market_def.get("keywords", []))}
Exclude: {", ".join(market_def.get("exclude_markets", []))}
""".strip()

    return f"""
{market_context}

Write IM-ready output using ONLY the facts provided.

Topic: {particular}

Row-specific instructions:
{row_prompt}

FACTS_JSON:
{facts_json}

Output EXACTLY:
DESC:
- Use bullets. Detailed but non-repetitive.
SOURCES:
- 4–8 URLs (one per line)

Rules:
- Do NOT add new facts.
- No markdown.
""".strip()


def evidence_gated_research_flat(
    company: str,
    particular: str,
    market_def: Dict[str, Any],
    client: OpenAI,
    facts_cache: Dict[Tuple[str, str], List[Dict[str, Any]]],
) -> Tuple[str, str]:
    cleaned_facts = fetch_facts_cached(company, particular, market_def, client, facts_cache)
    facts_json = json.dumps({"facts": cleaned_facts}, ensure_ascii=False)

    r2 = _responses_create(
        client,
        model=REWRITE_MODEL,
        input=build_im_prompt_flat(particular, facts_json, market_def),
        max_output_tokens=1400,
        temperature=0,
    )

    text = (r2.output_text or "").strip()
    if "SOURCES:" in text:
        before, after = text.split("SOURCES:", 1)
        desc = before.replace("DESC:", "").strip()
        sources_txt = after.strip()
    else:
        desc, sources_txt = text, ""

    desc = normalize_bullets(desc)
    evidence_urls = {canonical_url(f["evidence_url"]) for f in cleaned_facts}
    src_urls = [url for url in extract_urls(sources_txt) if url in evidence_urls]
    if not src_urls:
        src_urls = list(dict.fromkeys([f["evidence_url"] for f in cleaned_facts]))[:8]

    return clamp(desc, MAX_DESC_CHARS), clamp("\n".join(src_urls[:8]), MAX_SOURCES_CHARS)


# ---------------------------
# Rowwise bucket prompt (NEW schema: rows + sources only)
# ---------------------------
def build_im_prompt_rowwise(particular: str, facts_json: str, market_def: Dict[str, Any]) -> str:
    key = (particular or "").strip().lower()
    labels = BUCKET_ROW_SPECS.get(key, [])
    labels_json = json.dumps(labels, ensure_ascii=False)

    market_context = ""
    if key in {"industry landscape"} and market_def:
        market_context = f"""
MARKET DEFINITION (MANDATORY — do not deviate):
Primary market: {market_def.get("primary_market")}
Description: {market_def.get("market_description")}
Geography: {", ".join(market_def.get("include_geographies", []))}
Keywords: {", ".join(market_def.get("keywords", []))}
Exclude: {", ".join(market_def.get("exclude_markets", []))}
""".strip()

    return f"""
{market_context}

Write IM-ready output using ONLY the facts provided.

Topic: {particular}

FACTS_JSON:
{facts_json}

Return ONLY JSON in this exact schema:
{{
  "rows": [
    {{
      "label": "must be one of the required labels",
      "bullets": ["2–6 bullets specific to that label"]
    }}
  ],
  "sources": ["https://...", "https://..."]
}}

Rules:
- Required labels (must output ALL, in the same order; no new labels):
{labels_json}
- If a label has no information, write exactly one bullet: "Not publicly disclosed".
- Bullets must be factual, non-repetitive, and derived only from FACTS_JSON.
- sources must be 4–8 URLs, one per entry.
- No markdown.
""".strip()


def evidence_gated_research_rowwise(
    company: str,
    particular: str,
    market_def: Dict[str, Any],
    client: OpenAI,
    facts_cache: Dict[Tuple[str, str], List[Dict[str, Any]]],
) -> Tuple[List[Dict[str, Any]], List[str]]:
    cleaned_facts = fetch_facts_cached(company, particular, market_def, client, facts_cache)
    facts_json = json.dumps({"facts": cleaned_facts}, ensure_ascii=False)

    r2 = _responses_create(
        client,
        model=REWRITE_MODEL,
        input=build_im_prompt_rowwise(particular, facts_json, market_def),
        response_format={"type": "json_object"},
        max_output_tokens=1800,
        temperature=0,
    )

    raw = (r2.output_text or "").strip()
    out = extract_json_object(raw)
    if not out or "rows" not in out:
        schema_hint = """{ "rows":[{"label":"...","bullets":["..."]}], "sources":["https://..."] }"""
        out = repair_json_once(raw, schema_hint, client)

    rows = out.get("rows", [])
    sources = out.get("sources", [])

    if not isinstance(rows, list):
        rows = []
    if not isinstance(sources, list):
        sources = extract_urls(str(sources))

    # enforce label order
    key = (particular or "").strip().lower()
    required = BUCKET_ROW_SPECS.get(key, [])
    by_label: Dict[str, List[str]] = {}
    for r in rows:
        if not isinstance(r, dict):
            continue
        lab = str(r.get("label", "")).strip()
        bullets = r.get("bullets", [])
        if not isinstance(bullets, list):
            bullets = [str(bullets)]
        bullets = [str(x).strip() for x in bullets if str(x).strip()]
        if lab:
            by_label[lab] = bullets

    normalized_rows: List[Dict[str, Any]] = []
    for lab in required:
        bullets = by_label.get(lab, [])
        if not bullets:
            bullets = ["Not publicly disclosed"]
        normalized_rows.append({"label": lab, "bullets": bullets})

    evidence_urls = {canonical_url(f["evidence_url"]) for f in cleaned_facts}
    src_urls = [
        canonical_url(u)
        for u in sources
        if isinstance(u, str) and canonical_url(u) in evidence_urls
    ]
    if not src_urls:
        src_urls = list(dict.fromkeys([f["evidence_url"] for f in cleaned_facts]))[:8]

    return normalized_rows, src_urls[:8]


# ---------------------------
# Competitors / M&A / Buyers (JSON table outputs)
# ---------------------------
def infer_region_from_hq(hq: str) -> str:
    hq_l = (hq or "").lower()
    if "india" in hq_l:
        return "Domestic"
    return "Global"


def write_competitors(ws_comp, comp_json: Dict[str, Any]):
    row = ws_comp.max_row + 1

    bucket_map = [
        ("Direct competitor", "direct_competitors"),
        ("Adjacent peer", "adjacent_peers"),
    ]

    for label, key in bucket_map:
        items = comp_json.get(key, [])
        if not isinstance(items, list):
            continue

        for it in items:
            if not isinstance(it, dict):
                continue

            name = (it.get("name") or "").strip()
            hq = (it.get("hq") or "Not publicly disclosed").strip()
            desc = (it.get("description") or "").strip()
            sources = it.get("sources", [])
            if not isinstance(sources, list):
                sources = extract_urls(str(sources))
            sources = [canonical_url(s) for s in sources if isinstance(s, str) and s.startswith("http")]

            if not name or not sources:
                continue

            ws_comp.cell(row, COMP_COL_NAME).value = name
            ws_comp.cell(row, COMP_COL_TYPE).value = label
            ws_comp.cell(row, COMP_COL_REGION).value = infer_region_from_hq(hq)
            ws_comp.cell(row, COMP_COL_HQ).value = hq
            ws_comp.cell(row, COMP_COL_DESC).value = desc
            ws_comp.cell(row, COMP_COL_SOURCES).value = "\n".join(list(dict.fromkeys(sources))[:6])
            row += 1


def competitors_research_json(company: str, market_def: Dict[str, Any], client: OpenAI, facts_cache) -> Dict[str, Any]:
    particular = "competitors"
    cleaned_facts = fetch_facts_cached(company, particular, market_def, client, facts_cache)
    facts_json = json.dumps({"facts": cleaned_facts}, ensure_ascii=False)
    evidence_urls = {canonical_url(f["evidence_url"]) for f in cleaned_facts}

    market_context = ""
    if market_def:
        market_context = f"""
MARKET DEFINITION (MANDATORY — do not deviate):
Primary market: {market_def.get("primary_market")}
Description: {market_def.get("market_description")}
Geography: {", ".join(market_def.get("include_geographies", []))}
Keywords: {", ".join(market_def.get("keywords", []))}
Exclude: {", ".join(market_def.get("exclude_markets", []))}
""".strip()

    rewrite_prompt = f"""
{market_context}

Use ONLY the facts below to build a competitor list in JSON.

FACTS_JSON:
{facts_json}

{ROW_PROMPTS["competitors"]}
""".strip()

    r2 = _responses_create(
        client,
        model=REWRITE_MODEL,
        input=rewrite_prompt,
        response_format={"type": "json_object"},
        max_output_tokens=2200,
        temperature=0,
    )

    raw = (r2.output_text or "").strip()
    out = extract_json_object(raw)
    if not isinstance(out, dict) or "direct_competitors" not in out:
        schema_hint = """{ "direct_competitors":[{"name":"...","hq":"...","description":"...","sources":["https://..."]}], "adjacent_peers":[{"name":"...","hq":"...","description":"...","sources":["https://..."]}] }"""
        out = repair_json_once(raw, schema_hint, client)

    def clean_list(arr):
        cleaned = []
        if not isinstance(arr, list):
            return cleaned
        for x in arr:
            if not isinstance(x, dict):
                continue
            name = str(x.get("name", "")).strip()
            hq = str(x.get("hq", "")).strip() or "Not publicly disclosed"
            desc = str(x.get("description", "")).strip()
            sources = x.get("sources", [])
            if not isinstance(sources, list):
                sources = extract_urls(str(sources))
            sources = [
                canonical_url(s)
                for s in sources
                if isinstance(s, str) and canonical_url(s) in evidence_urls
            ]
            if name and sources:
                cleaned.append({"name": name, "hq": hq, "description": desc, "sources": list(dict.fromkeys(sources))[:6]})
        return cleaned

    result = {
        "direct_competitors": clean_list(out.get("direct_competitors", []))[:12],
        "adjacent_peers": clean_list(out.get("adjacent_peers", []))[:12],
    }
    return result


def write_ma(ws_ma, ma_json: Dict[str, Any]):
    row = ws_ma.max_row + 1
    for d in ma_json.get("deals", []):
        ws_ma.cell(row, 1).value = d.get("acquirer")
        ws_ma.cell(row, 2).value = d.get("target")
        ws_ma.cell(row, 3).value = d.get("year")
        ws_ma.cell(row, 4).value = d.get("deal_value")
        ws_ma.cell(row, 5).value = d.get("rationale")
        ws_ma.cell(row, 6).value = d.get("source")
        row += 1


def ma_research_json(company: str, market_def: Dict[str, Any], client: OpenAI, facts_cache) -> Dict[str, Any]:
    particular = "m&a landscape"
    cleaned_facts = fetch_facts_cached(company, particular, market_def, client, facts_cache)
    facts_json = json.dumps({"facts": cleaned_facts}, ensure_ascii=False)
    evidence_urls = {canonical_url(f["evidence_url"]) for f in cleaned_facts}

    market_context = ""
    if market_def:
        market_context = f"""
MARKET DEFINITION (MANDATORY — do not deviate):
Primary market: {market_def.get("primary_market")}
Description: {market_def.get("market_description")}
Geography: {", ".join(market_def.get("include_geographies", []))}
Keywords: {", ".join(market_def.get("keywords", []))}
Exclude: {", ".join(market_def.get("exclude_markets", []))}
""".strip()

    rewrite_prompt = f"""
{market_context}

Using ONLY FACTS_JSON, produce a structured deal list for Excel.

FACTS_JSON:
{facts_json}

Return ONLY JSON:
{{
  "deals": [
    {{
      "acquirer": "string",
      "target": "string",
      "year": "YYYY or Not disclosed",
      "deal_value": "value or Not disclosed",
      "rationale": "1 line",
      "source": "https://..."
    }}
  ]
}}

Rules:
- Each deal must have a source URL.
- No placeholders; if missing, use "Not disclosed".
""".strip()

    r2 = _responses_create(
        client,
        model=REWRITE_MODEL,
        input=rewrite_prompt,
        response_format={"type": "json_object"},
        max_output_tokens=2200,
        temperature=0,
    )

    raw = (r2.output_text or "").strip()
    out = extract_json_object(raw)
    if not isinstance(out, dict) or "deals" not in out:
        schema_hint = """{ "deals":[{"acquirer":"...","target":"...","year":"YYYY","deal_value":"...","rationale":"...","source":"https://..."}] }"""
        out = repair_json_once(raw, schema_hint, client)

    deals = out.get("deals", [])
    cleaned = []
    if isinstance(deals, list):
        for d in deals:
            if not isinstance(d, dict):
                continue
            src = canonical_url(str(d.get("source", "")).strip())
            if src not in evidence_urls:
                continue
            cleaned.append({
                "acquirer": str(d.get("acquirer", "")).strip(),
                "target": str(d.get("target", "")).strip(),
                "year": str(d.get("year", "")).strip() or "Not disclosed",
                "deal_value": str(d.get("deal_value", "")).strip() or "Not disclosed",
                "rationale": str(d.get("rationale", "")).strip(),
                "source": src,
            })

    # sort newest first when possible
    def _year_num(y: str) -> int:
        try:
            return int(re.findall(r"\d{4}", str(y))[0])
        except Exception:
            return -1

    cleaned.sort(key=lambda d: _year_num(d.get("year", "")), reverse=True)
    return {"deals": cleaned[:15]}


def write_buyers(ws_buyers, buyers_json: Dict[str, Any]):
    row = ws_buyers.max_row + 1
    for b in buyers_json.get("buyers", []):
        ws_buyers.cell(row, 1).value = b.get("buyer")
        ws_buyers.cell(row, 2).value = b.get("category")
        ws_buyers.cell(row, 3).value = b.get("hq")
        ws_buyers.cell(row, 4).value = b.get("rationale")
        ws_buyers.cell(row, 5).value = b.get("frictions")
        ws_buyers.cell(row, 6).value = b.get("prior_acquisitions")
        ws_buyers.cell(row, 7).value = b.get("source")
        row += 1


def buyers_research_json(company: str, market_def: Dict[str, Any], client: OpenAI, facts_cache) -> Dict[str, Any]:
    particular = "prospective buyers"
    cleaned_facts = fetch_facts_cached(company, particular, market_def, client, facts_cache)
    facts_json = json.dumps({"facts": cleaned_facts}, ensure_ascii=False)
    evidence_urls = {canonical_url(f["evidence_url"]) for f in cleaned_facts}

    market_context = ""
    if market_def:
        market_context = f"""
MARKET DEFINITION (MANDATORY — do not deviate):
Primary market: {market_def.get("primary_market")}
Description: {market_def.get("market_description")}
Geography: {", ".join(market_def.get("include_geographies", []))}
Keywords: {", ".join(market_def.get("keywords", []))}
Exclude: {", ".join(market_def.get("exclude_markets", []))}
""".strip()

    rewrite_prompt = f"""
{market_context}

Using ONLY FACTS_JSON, produce a structured buyer universe for Excel.

FACTS_JSON:
{facts_json}

Return ONLY JSON:
{{
  "buyers": [
    {{
      "buyer": "string",
      "category": "Strategic | Competitor | Infra/PE Platform | Adjacent Strategic | Other",
      "hq": "string or Not disclosed",
      "rationale": "1 line",
      "frictions": "1 line",
      "prior_acquisitions": "string or Not disclosed",
      "source": "https://..."
    }}
  ]
}}

Rules:
- Each buyer must have a source URL.
- No invented acquisitions; use "Not disclosed".
""".strip()

    r2 = _responses_create(
        client,
        model=REWRITE_MODEL,
        input=rewrite_prompt,
        response_format={"type": "json_object"},
        max_output_tokens=2400,
        temperature=0,
    )

    raw = (r2.output_text or "").strip()
    out = extract_json_object(raw)
    if not isinstance(out, dict) or "buyers" not in out:
        schema_hint = """{ "buyers":[{"buyer":"...","category":"...","hq":"...","rationale":"...","frictions":"...","prior_acquisitions":"...","source":"https://..."}] }"""
        out = repair_json_once(raw, schema_hint, client)

    buyers = out.get("buyers", [])
    cleaned = []
    if isinstance(buyers, list):
        for b in buyers:
            if not isinstance(b, dict):
                continue
            src = canonical_url(str(b.get("source", "")).strip())
            if src not in evidence_urls:
                continue
            cleaned.append({
                "buyer": str(b.get("buyer", "")).strip(),
                "category": str(b.get("category", "")).strip(),
                "hq": str(b.get("hq", "")).strip() or "Not disclosed",
                "rationale": str(b.get("rationale", "")).strip(),
                "frictions": str(b.get("frictions", "")).strip(),
                "prior_acquisitions": str(b.get("prior_acquisitions", "")).strip() or "Not disclosed",
                "source": src,
            })

    # basic dedupe by name
    def _norm(s: str) -> str:
        return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()

    seen = set()
    deduped = []
    for b in cleaned:
        k = _norm(b.get("buyer", ""))
        if k and k not in seen:
            seen.add(k)
            deduped.append(b)

    return {"buyers": deduped[:25]}


# =========================
# pipeline_core_v2.py
# PART 3 / 3 — Main generate_profile_excel_bytes (NEW main-sheet layout)
# =========================

def generate_profile_excel_bytes(
    company_name: str,
    encrypted_id: str,
    openai_api_key: str,
    privatecircle_token: str,
    progress_cb=None,
    cin: str = "",
    mergermarket_export_name: str = "",
    mergermarket_export_bytes: Optional[bytes] = None,
    transaction_scope: str = "Domestic and global",
    mergermarket_email: str = "",
    mergermarket_password: str = "",
    transaction_start_date: Optional[date] = None,
    transaction_end_date: Optional[date] = None,
    mergermarket_max_deals_per_query: int = 100,
) -> bytes:
    """
    Front-end contract unchanged:
      inputs: company_name, encrypted_id, openai_api_key, privatecircle_token
      output: Excel bytes
    """

    company_name = (company_name or "").strip()
    encrypted_id = (encrypted_id or "").strip()
    openai_api_key = (openai_api_key or "").strip()
    privatecircle_token = (privatecircle_token or "").strip()
    cin = (cin or "").strip()
    mergermarket_export_name = (mergermarket_export_name or "").strip()
    transaction_scope = (transaction_scope or "Domestic and global").strip()
    mergermarket_email = (mergermarket_email or "").strip()
    mergermarket_password = mergermarket_password or ""
    

    if not company_name:
        raise ValueError("Company name is required.")
    if not encrypted_id:
        raise ValueError("PrivateCircle company id is required.")
    if not openai_api_key:
        raise ValueError("OpenAI API key is required.")
    if not privatecircle_token:
        raise ValueError("PrivateCircle token is required.")

    def cb(msg: str, pct: int):
        if progress_cb:
            try:
                progress_cb(msg, pct)
            except Exception:
                pass

    cb("Loading template...", 2)
    wb = load_template_workbook()

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Sheet '{SHEET_NAME}' not found in template. Available: {wb.sheetnames}")
    ws = wb[SHEET_NAME]

    # Company name cell (as in your original design)
    ws[COMPANY_NAME_CELL].value = company_name

    # Ensure Description column (E) has enough width
    ws.column_dimensions['E'].width = 70
    ws.column_dimensions['F'].width = 40
    # Ensure sheets exist and clear table sheets
    ws_comp = ensure_competitors_sheet(wb)
    ws_ma = ensure_ma_sheet(wb)
    ws_buyers = ensure_buyers_sheet(wb)
    ensure_financials_sheet(wb)
    ws_credit = ensure_credit_rating_sheet(wb)
    ws_sources = ensure_source_dossier_sheet(wb)

    clear_table_sheet_keep_header(ws_comp)
    clear_table_sheet_keep_header(ws_ma)
    clear_table_sheet_keep_header(ws_buyers)
    clear_table_sheet_keep_header(ws_credit)
    clear_table_sheet_keep_header(ws_sources)

    cb("Pulling financials from PrivateCircle...", 8)
    try:
        write_financials_sheet(wb, encrypted_id, privatecircle_token)
    except Exception as e:
        ws_fin = ensure_financials_sheet(wb)
        ws_fin.cell(1, 1).value = f"Financials not populated due to error: {e}"

    cb("Discovering official website and primary sources...", 12)
    dossier: Dict[str, Any] = {}
    credit_reports: List[Dict[str, str]] = []
    orchestrator = None
    try:
        orchestrator = ResearchOrchestrator(openai_api_key, model=RESEARCH_MODEL)
        research_result = orchestrator.run_company_dossier(company_name, cin=cin)
        dossier = research_result.get("dossier", {})
        credit_reports = research_result.get("credit_reports", [])
        write_source_dossier(ws_sources, dossier)
        write_credit_reports(ws_credit, credit_reports)
        if not credit_reports:
            ws_credit.cell(2, 1).value = "No exact official rating-agency report found"
            ws_credit.cell(2, 8).value = "No verified match was returned for this legal entity."
    except Exception:
        # Specialist discovery is additive. Preserve the core profile if the
        # web-research stage is temporarily unavailable.
        ws_sources.cell(2, 1).value = "research_error"
        ws_sources.cell(2, 6).value = "Source discovery was temporarily unavailable."
        ws_credit.cell(2, 1).value = "Research temporarily unavailable"

    format_source_dossier_sheet(ws_sources)
    format_credit_rating_sheet(ws_credit)

    client = OpenAI(api_key=openai_api_key, timeout=120.0, max_retries=2)
    facts_cache: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}

    cb("Building market definition (web search)...", 22)
    market_def = get_market_definition(company_name, client=client)
    market_def["legal_name"] = dossier.get("legal_name") or company_name
    market_def["official_website"] = dossier.get("official_website") or ""
    market_def["verified_sources"] = [
        source.get("url")
        for source in dossier.get("sources", [])
        if isinstance(source, dict) and source.get("url")
    ]

    has_uploaded_mergermarket_export = bool(mergermarket_export_bytes and mergermarket_export_name)
    wants_automatic_mergermarket = bool(mergermarket_email and mergermarket_password)
    has_mergermarket_data = False
    if has_uploaded_mergermarket_export:
        cb("Importing and screening Mergermarket transactions...", 27)
        imported = parse_mergermarket_export(mergermarket_export_name, mergermarket_export_bytes)
        search_queries: List[str] = []
        retrieval_mode = "Authorised Mergermarket export"
        has_mergermarket_data = True
    elif wants_automatic_mergermarket:
        cb("Signing in and screening Mergermarket Deals...", 27)
        if orchestrator is None:
            orchestrator = ResearchOrchestrator(openai_api_key, model=RESEARCH_MODEL)
        try:
            search_queries = orchestrator.build_mergermarket_search_plan(company_name, market_def)
        except Exception:
            search_queries = []
        if not search_queries:
            keyword_fallback = [
                str(value).strip()
                for value in market_def.get("keywords", [])
                if str(value).strip()
            ][:8]
            market_label = str(market_def.get("primary_market") or company_name).strip()
            search_queries = [f"Target companies in {market_label}: {', '.join(keyword_fallback)}"]
        end_date = transaction_end_date or date.today()
        try:
            default_start = end_date.replace(year=end_date.year - 10)
        except ValueError:
            default_start = end_date.replace(year=end_date.year - 10, day=28)
        start_date = transaction_start_date or default_start
        browser_result = fetch_mergermarket_transactions(
            MergermarketCredentials(email=mergermarket_email, password=mergermarket_password),
            MergermarketSearchPlan(
                queries=search_queries,
                start_date=start_date,
                end_date=end_date,
                geography_scope=transaction_scope,
                max_deals_per_query=max(50, min(500, int(mergermarket_max_deals_per_query))),
            ),
        )
        imported = browser_result.imported
        imported.warnings.extend(browser_result.warnings)
        search_queries = browser_result.executed_queries
        retrieval_mode = "Automatic authorised Mergermarket retrieval"
        has_mergermarket_data = True

    if has_mergermarket_data:
        try:
            if orchestrator is None:
                orchestrator = ResearchOrchestrator(openai_api_key, model=RESEARCH_MODEL)
            orchestrator.classify_transactions(company_name, market_def, imported.deals)
        except Exception:
            imported.warnings.append("Automated relevance screening was unavailable; candidates require manual review.")
        write_transaction_comps_workbook(
            wb,
            imported,
            company_name,
            transaction_scope,
            search_queries=search_queries,
            retrieval_mode=retrieval_mode,
        )

    cb("Filling main sheet...", 18)

    # We process based on Column A (section headings).
    # For bucketed sections we will insert rows and merge A cells.
    r = START_ROW
    safety = 0

    # keys we will handle specially
    bucketed_keys = set(BUCKET_ROW_SPECS.keys())
    special_keys = {
        "about the company",
        "competitors",
        "m&a landscape",
        "prospective buyers",
        "financials",
        "financial overview",
        "financial snapshot",
        "company financials",
        "financial performance",
    }

    while r <= ws.max_row and safety < 8000:
        safety += 1

        section = ws.cell(r, MAIN_COL_SECTION).value
        if not section or not str(section).strip():
            r += 1
            continue

        section = str(section).strip()
        key = section.lower()

        # clear row outputs (answers + sources only; preserve template labels in column D)
        ws.cell(r, MAIN_COL_ANSWER).value = ""
        ws.cell(r, MAIN_COL_SOURCES).value = ""

        # Financials pointer row
        if key in {"financials", "financial overview", "financial snapshot", "company financials", "financial performance"}:
            ws.cell(r, MAIN_COL_ANSWER).value = "See 'Financials' sheet"
            r += 1
            continue

        # Competitors sheet
        if key == "competitors":
            cb("Researching competitors (web search)...", 40)
            comp_json = competitors_research_json(company_name, market_def, client=client, facts_cache=facts_cache)
            write_competitors(ws_comp, comp_json)
            format_competitors_sheet(ws_comp)
            ws.cell(r, MAIN_COL_ANSWER).value = "See 'Competitors' sheet"
            time.sleep(SLEEP_BETWEEN_CALLS_SEC)
            r += 1
            continue

        # M&A sheet
        if key == "m&a landscape":
            cb("Researching M&A landscape (web search)...", 55)
            ma_json = ma_research_json(company_name, market_def, client=client, facts_cache=facts_cache)
            write_ma(ws_ma, ma_json)
            format_ma_sheet(ws_ma)
            if has_mergermarket_data:
                ws.cell(r, MAIN_COL_ANSWER).value = (
                    "See 'Final Comps', 'Adjacent', 'Excluded', 'QA', and complementary public-web 'M&A Landscape' sheets"
                )
            else:
                ws.cell(r, MAIN_COL_ANSWER).value = "See 'M&A Landscape' sheet"
            time.sleep(SLEEP_BETWEEN_CALLS_SEC)
            r += 1
            continue

        # Buyers sheet
        if key == "prospective buyers":
            cb("Researching prospective buyers (web search)...", 70)
            buyers_json = buyers_research_json(company_name, market_def, client=client, facts_cache=facts_cache)
            write_buyers(ws_buyers, buyers_json)
            format_buyers_sheet(ws_buyers)
            ws.cell(r, MAIN_COL_ANSWER).value = "See 'Prospective Buyers' sheet"
            time.sleep(SLEEP_BETWEEN_CALLS_SEC)
            r += 1
            continue

        # About the company (flat single-row)
        if key == "about the company":
            cb("Researching: About the company", 30)
            desc, sources = evidence_gated_research_flat(
                company_name, section, market_def, client=client, facts_cache=facts_cache
            )
            ws.cell(r, MAIN_COL_ANSWER).value = desc
            ws.cell(r, MAIN_COL_SOURCES).value = sources
            time.sleep(SLEEP_BETWEEN_CALLS_SEC)
            r += 1
            continue

        # Bucketed matrix sections
        if key in bucketed_keys:
            cb(f"Researching (bucketed matrix): {section}", 30)
            rows, src_urls = evidence_gated_research_rowwise(
                company_name, section, market_def, client=client, facts_cache=facts_cache
            )
            used = write_bucketed_matrix(
                ws=ws,
                start_row=r,
                section_title=section,
                rows=rows,
                sources=src_urls,
            )
            time.sleep(SLEEP_BETWEEN_CALLS_SEC)
            r += used
            continue

        # Other rows (flat)
        cb(f"Researching: {section}", 30)
        desc, sources = evidence_gated_research_flat(
            company_name, section, market_def, client=client, facts_cache=facts_cache
        )
        ws.cell(r, MAIN_COL_ANSWER).value = desc
        ws.cell(r, MAIN_COL_SOURCES).value = sources
        time.sleep(SLEEP_BETWEEN_CALLS_SEC)
        r += 1

    # Final formatting on auxiliary sheets
    write_evidence_register(wb, facts_cache)
    if COMPETITORS_SHEET in wb.sheetnames:
        format_competitors_sheet(wb[COMPETITORS_SHEET])
    if MA_SHEET in wb.sheetnames:
        format_ma_sheet(wb[MA_SHEET])
    if BUYERS_SHEET in wb.sheetnames:
        format_buyers_sheet(wb[BUYERS_SHEET])
    if CREDIT_RATING_SHEET in wb.sheetnames:
        format_credit_rating_sheet(wb[CREDIT_RATING_SHEET])
    if SOURCE_DOSSIER_SHEET in wb.sheetnames:
        format_source_dossier_sheet(wb[SOURCE_DOSSIER_SHEET])

    cb("Preparing download...", 98)
    bio = io.BytesIO()
    wb.save(bio)
    cb("Done.", 100)
    return bio.getvalue()
