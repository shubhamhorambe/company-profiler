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
from typing import Optional, Dict, Any, Tuple, List
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openai import OpenAI


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

RESEARCH_MODEL = "gpt-4o"
REWRITE_MODEL = "gpt-4o-mini"

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
            if r.status_code == 429:
                time.sleep(1.5 * (attempt + 1))
                continue
            if r.status_code >= 400:
                raise RuntimeError(f"PrivateCircle HTTP {r.status_code}: {r.text[:800]}")
            return r.json()
        except Exception as e:
            last_err = e
            time.sleep(0.6 * (attempt + 1))
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
# Financials (PrivateCircle) — your logic preserved
# ---------------------------
def to_inr_mn(value):
    try:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return round(value / 1_000_000, 2)
        v = float(str(value).replace(",", ""))
        return round(v / 1_000_000, 2)
    except Exception:
        return None


def pull_financial_statements(encrypted_id: str, token: str, no_of_years: int = 5, xbrl_type_tag: str = "consolidated") -> Dict[str, Any]:
    ov = pc_get_company_endpoint(encrypted_id, "overview", token)

    inc = pc_get_company_endpoint(
        encrypted_id,
        "income-statement",
        token,
        params={
            "no_of_years": no_of_years,
            "source": "mca",
            "type": "detailed",
            "xbrl_type_tag": xbrl_type_tag,
            "field_list": [
                "fiscal_year",
                "revenue_from_operations",
                "total_revenue",
                "total_sales_wo_other_income",
                "other_income",
                "total_export_sales",
                "export_goods_manf_sales",
                "export_goods_traded_sales",
                "export_services_sales",
                "gross_profit",
                "operating_ebitda",
                "operating_ebit",
                "pat",
            ],
        },
    )

    bs = pc_get_company_endpoint(
        encrypted_id,
        "balance-sheet",
        token,
        params={
            "no_of_years": no_of_years,
            "source": "mca",
            "type": "detailed",
            "xbrl_type_tag": xbrl_type_tag,
            "field_list": ["fiscal_year", "payable_outstanding_days", "sales_outstanding_days", "inventory_outstanding_days"],
        },
    )

    ratios = pc_get_company_endpoint(
        encrypted_id,
        "basic-financial-ratios",
        token,
        params={
            "no_of_years": no_of_years,
            "source": "mca",
            "xbrl_type_tag": xbrl_type_tag,
            "field_list": ["fiscal_year", "roce_percent"],
        },
    )

    return {"overview": ov, "income_statement": inc, "balance_sheet": bs, "basic_financial_ratios": ratios}


def apply_financials_format(ws, max_row: int, n_years: int):
    title_font = Font(bold=True, size=14)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    section_fill = PatternFill("solid", fgColor="D9E1F2")
    bold = Font(bold=True)

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"].font = title_font
    ws.column_dimensions["A"].width = 34
    for j in range(n_years):
        ws.column_dimensions[get_column_letter(2 + j)].width = 16
    ws.column_dimensions["B"].width = 18

    ws.freeze_panes = "B7"

    for r in range(1, max_row + 1):
        for c in range(1, 2 + max(n_years, 1)):
            cell = ws.cell(r, c)
            if cell.value is None:
                continue
            cell.border = border
            if c == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(1, max_row + 1):
        if ws.cell(r, 1).value == "Metric":
            for c in range(1, 2 + max(n_years, 1)):
                cell = ws.cell(r, c)
                cell.fill = header_fill
                cell.font = header_font

    for r in range(1, max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v in {
            "Income Statement (Last 5 years)",
            "Balance Sheet (Working Capital Days)",
            "Basic Financial Ratios",
            "3Y Summary Metrics",
        }:
            for c in range(1, 2 + max(n_years, 1)):
                ws.cell(r, c).fill = section_fill
            ws.cell(r, 1).font = bold


def write_financials_sheet(wb: openpyxl.Workbook, encrypted_id: str, token: str):
    ws = ensure_financials_sheet(wb)

    fin = pull_financial_statements(encrypted_id, token, no_of_years=5, xbrl_type_tag="consolidated")
    inc_items = fin.get("income_statement", {}).get("items") or []
    if not inc_items:
        fin = pull_financial_statements(encrypted_id, token, no_of_years=5, xbrl_type_tag="standalone")
        inc_items = fin.get("income_statement", {}).get("items") or []

    bs_items = fin.get("balance_sheet", {}).get("items") or []
    ratios_items = fin.get("basic_financial_ratios", {}).get("items") or []
    ov_items = fin.get("overview", {}).get("items") or []
    ov = ov_items[0] if ov_items else {}

    if not inc_items:
        ws.cell(1, 1).value = "No Income Statement data available from PrivateCircle."
        return

    years = sorted({int(x["fiscal_year"]) for x in inc_items if x.get("fiscal_year")}, reverse=True)[:5]
    years = sorted(years)

    inc_by_year = {int(x["fiscal_year"]): x for x in inc_items if x.get("fiscal_year")}
    bs_by_year = {int(x["fiscal_year"]): x for x in bs_items if x.get("fiscal_year")}
    ratios_by_year = {int(x["fiscal_year"]): x for x in ratios_items if x.get("fiscal_year")}

    def num(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().lower()
        if s in {"", "na", "nm", "none"}:
            return None
        try:
            return float(str(v).replace(",", ""))
        except Exception:
            return None

    def exports_val(row: Dict[str, Any]) -> float:
        v = num(row.get("total_export_sales"))
        if v is not None:
            return v
        a = num(row.get("export_goods_manf_sales")) or 0.0
        b = num(row.get("export_goods_traded_sales")) or 0.0
        c = num(row.get("export_services_sales")) or 0.0
        return a + b + c

    def total_income_val(row):
        v = num(row.get("total_revenue"))
        if v is not None:
            return v
        ops = num(row.get("revenue_from_operations"))
        if ops is None:
            ops = num(row.get("total_sales_wo_other_income"))
        other = num(row.get("other_income")) or 0.0
        if ops is not None:
            return ops + other
        return None

    ws.cell(1, 1).value = "Financials (Source: PrivateCircle)"
    ws.cell(2, 1).value = "Company"
    ws.cell(2, 2).value = ov.get("name") or "Not disclosed"
    ws.cell(3, 1).value = "Note: API returns figures as reported (see PrivateCircle notes)."

    r = 5
    ws.cell(r, 1).value = "Income Statement (Last 5 years)"
    r += 1

    ws.cell(r, 1).value = "Metric"
    for j, y in enumerate(years):
        ws.cell(r, 2 + j).value = f"FY{y}"
    r += 1

    start_col = 2
    n = len(years)
    row_map = {}

    def write_value_row(label: str, getter):
        nonlocal r
        ws.cell(r, 1).value = label
        for j, y in enumerate(years):
            row = inc_by_year.get(y, {})
            v = getter(row)
            ws.cell(r, start_col + j).value = v
        row_map[label] = r
        r += 1

    def write_percent_row(label: str, numerator_label: str, denom_label: str):
        nonlocal r
        ws.cell(r, 1).value = label
        nr = row_map[numerator_label]
        dr = row_map[denom_label]
        for j in range(n):
            col = get_column_letter(start_col + j)
            ws.cell(r, start_col + j).value = f'=IFERROR({col}{nr}/{col}{dr},"")'
            ws.cell(r, start_col + j).number_format = "0.00%"
        row_map[label] = r
        r += 1

    write_value_row("Revenue from Operations", lambda row: to_inr_mn(num(row.get("revenue_from_operations"))))
    write_value_row("Total Revenue", lambda row: to_inr_mn(total_income_val(row)))
    write_value_row("Exports", lambda row: to_inr_mn(exports_val(row)))
    write_percent_row("Exports %", "Exports", "Revenue from Operations")

    write_value_row("Gross Profit", lambda row: to_inr_mn(num(row.get("gross_profit"))))
    write_percent_row("GP %", "Gross Profit", "Revenue from Operations")

    write_value_row("Operating EBITDA", lambda row: to_inr_mn(num(row.get("operating_ebitda"))))
    write_percent_row("Operating EBITDA%", "Operating EBITDA", "Revenue from Operations")

    write_value_row("Operating EBIT", lambda row: to_inr_mn(num(row.get("operating_ebit"))))
    write_percent_row("Operating EBIT%", "Operating EBIT", "Revenue from Operations")

    write_value_row("PAT", lambda row: to_inr_mn(num(row.get("pat"))))
    write_percent_row("PAT %", "PAT", "Revenue from Operations")

    amount_rows = {"Revenue from Operations", "Total Revenue", "Exports", "Gross Profit", "Operating EBITDA", "Operating EBIT", "PAT"}
    for label in amount_rows:
        rr = row_map[label]
        for j in range(n):
            ws.cell(rr, start_col + j).number_format = "#,##0"

    r += 1

    ws.cell(r, 1).value = "Balance Sheet (Working Capital Days)"
    r += 1
    ws.cell(r, 1).value = "Metric"
    for j, y in enumerate(years):
        ws.cell(r, 2 + j).value = f"FY{y}"
    r += 1

    def write_bs_row(label: str, field: str):
        nonlocal r
        ws.cell(r, 1).value = label
        for j, y in enumerate(years):
            ws.cell(r, 2 + j).value = num(bs_by_year.get(y, {}).get(field))
            ws.cell(r, 2 + j).number_format = "0.0"
        row_map[label] = r
        r += 1

    write_bs_row("Days Trade payables", "payable_outstanding_days")
    write_bs_row("Days Sales outstanding", "sales_outstanding_days")
    write_bs_row("Days Inventory outstanding", "inventory_outstanding_days")

    ws.cell(r, 1).value = "Net cash conversion cycle"
    pay_r = row_map["Days Trade payables"]
    so_r = row_map["Days Sales outstanding"]
    inv_r = row_map["Days Inventory outstanding"]
    for j in range(n):
        col = get_column_letter(2 + j)
        ws.cell(r, 2 + j).value = f'=IFERROR({col}{inv_r}+{col}{so_r}-{col}{pay_r},"")'
        ws.cell(r, 2 + j).number_format = "0.0"
    row_map["Net cash conversion cycle"] = r
    r += 2

    ws.cell(r, 1).value = "Basic Financial Ratios"
    r += 1
    ws.cell(r, 1).value = "Metric"
    for j, y in enumerate(years):
        ws.cell(r, 2 + j).value = f"FY{y}"
    r += 1

    ws.cell(r, 1).value = "RoCE"
    roce_row = r
    for j, y in enumerate(years):
        v = num((ratios_by_year.get(y) or {}).get("roce_percent"))
        if v is None:
            ws.cell(r, 2 + j).value = ""
        else:
            ws.cell(r, 2 + j).value = v / 100.0
            ws.cell(r, 2 + j).number_format = "0.00%"
    r += 2

    ws.cell(r, 1).value = "3Y Summary Metrics"
    r += 1
    ws.cell(r, 1).value = "Metric"
    ws.cell(r, 2).value = "Value"
    r += 1

    if n >= 4:
        last_idx = n - 1
        start3_idx = n - 4

        rev_r = row_map["Revenue from Operations"]
        ebitda_r = row_map["Operating EBITDA"]

        rev_end = f"{get_column_letter(start_col + last_idx)}{rev_r}"
        rev_start = f"{get_column_letter(start_col + start3_idx)}{rev_r}"
        e_end = f"{get_column_letter(start_col + last_idx)}{ebitda_r}"
        e_start = f"{get_column_letter(start_col + start3_idx)}{ebitda_r}"

        ws.cell(r, 1).value = "3Y Revenue CAGR"
        ws.cell(r, 2).value = f'=IFERROR(POWER({rev_end}/{rev_start},1/3)-1,"")'
        ws.cell(r, 2).number_format = "0.00%"
        r += 1

        ws.cell(r, 1).value = "3Y EBITDA CAGR"
        ws.cell(r, 2).value = f'=IFERROR(POWER({e_end}/{e_start},1/3)-1,"")'
        ws.cell(r, 2).number_format = "0.00%"
        r += 1

        c1 = get_column_letter(2 + (n - 3))
        c2 = get_column_letter(2 + (n - 2))
        c3 = get_column_letter(2 + (n - 1))
        ws.cell(r, 1).value = "3Y Average RoCE"
        ws.cell(r, 2).value = f'=IFERROR(AVERAGE({c1}{roce_row},{c2}{roce_row},{c3}{roce_row}),"")'
        ws.cell(r, 2).number_format = "0.00%"
        r += 1

    apply_financials_format(ws, max_row=r, n_years=n)


# ---------------------------
# OpenAI response helper (supports older SDKs)
# ---------------------------
def _responses_create(client: OpenAI, **kwargs):
    """
    Tries response_format if supported; falls back cleanly.
    """
    try:
        return client.responses.create(**kwargs)
    except TypeError:
        kwargs.pop("response_format", None)
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

    return f"""
Use web search and collect ONLY verifiable facts.

Company: {company}
Topic: {particular}

{market_context}

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
    src_urls = extract_urls(sources_txt)
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

    src_urls = [canonical_url(u) for u in sources if isinstance(u, str) and u.startswith("http")]
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
            sources = [canonical_url(s) for s in sources if isinstance(s, str) and s.startswith("http")]
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
            if not src.startswith("http"):
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
            if not src.startswith("http"):
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

    clear_table_sheet_keep_header(ws_comp)
    clear_table_sheet_keep_header(ws_ma)
    clear_table_sheet_keep_header(ws_buyers)

    cb("Pulling financials from PrivateCircle...", 8)
    try:
        write_financials_sheet(wb, encrypted_id, privatecircle_token)
    except Exception as e:
        ws_fin = ensure_financials_sheet(wb)
        ws_fin.cell(1, 1).value = f"Financials not populated due to error: {e}"

    client = OpenAI(api_key=openai_api_key)
    facts_cache: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}

    cb("Building market definition (web search)...", 15)
    market_def = get_market_definition(company_name, client=client)

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
    if COMPETITORS_SHEET in wb.sheetnames:
        format_competitors_sheet(wb[COMPETITORS_SHEET])
    if MA_SHEET in wb.sheetnames:
        format_ma_sheet(wb[MA_SHEET])
    if BUYERS_SHEET in wb.sheetnames:
        format_buyers_sheet(wb[BUYERS_SHEET])

    cb("Preparing download...", 98)
    bio = io.BytesIO()
    wb.save(bio)
    cb("Done.", 100)
    return bio.getvalue()