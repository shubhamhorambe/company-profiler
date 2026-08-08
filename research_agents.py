"""Specialist web-research agents used by the company profiler.

The agents in this module are deliberately narrow: they discover sources and
return structured evidence. They do not write the workbook or make business
decisions. Keeping those responsibilities deterministic makes the generated
profile easier to audit and test.
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, Iterable, List
from urllib.parse import urlsplit, urlunsplit

from openai import OpenAI


DEFAULT_RESEARCH_MODEL = os.getenv("OPENAI_RESEARCH_MODEL", "gpt-5.6-terra")

RATING_AGENCY_DOMAINS = {
    "acuite.in",
    "brickworkratings.com",
    "careedge.in",
    "careratings.com",
    "crisil.com",
    "crisilratings.com",
    "icra.in",
    "indiaratings.co.in",
    "infomerics.com",
}

SOURCE_TYPES = {
    "official_website",
    "credit_rating",
    "annual_report",
    "investor_presentation",
    "regulatory_filing",
    "registry",
    "reputable_news",
    "industry_source",
    "other",
}


def _extract_json_object(text: str) -> Dict[str, Any]:
    text = (text or "").strip()
    if not text:
        return {}
    try:
        value = json.loads(text)
        return value if isinstance(value, dict) else {}
    except json.JSONDecodeError:
        pass

    start = text.find("{")
    if start < 0:
        return {}
    depth = 0
    quoted = False
    escaped = False
    for index in range(start, len(text)):
        char = text[index]
        if quoted:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                quoted = False
            continue
        if char == '"':
            quoted = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                try:
                    value = json.loads(text[start : index + 1])
                    return value if isinstance(value, dict) else {}
                except json.JSONDecodeError:
                    return {}
    return {}


def normalize_url(value: str) -> str:
    """Return a safe, canonical http(s) URL or an empty string."""
    try:
        parts = urlsplit((value or "").strip())
    except ValueError:
        return ""
    if parts.scheme.lower() not in {"http", "https"} or not parts.hostname:
        return ""
    host = parts.hostname.lower().rstrip(".")
    if host in {"localhost", "127.0.0.1", "0.0.0.0"} or host.endswith(".local"):
        return ""
    netloc = host
    if parts.port:
        netloc = f"{host}:{parts.port}"
    return urlunsplit((parts.scheme.lower(), netloc, parts.path or "/", parts.query, ""))


def _host_matches(url: str, domains: Iterable[str]) -> bool:
    try:
        host = (urlsplit(url).hostname or "").lower()
    except ValueError:
        return False
    return any(host == domain or host.endswith(f".{domain}") for domain in domains)


def clean_source_dossier(payload: Dict[str, Any]) -> Dict[str, Any]:
    official_website = normalize_url(str(payload.get("official_website", "")))
    sources: List[Dict[str, str]] = []
    seen = set()
    for item in payload.get("sources", []) if isinstance(payload.get("sources"), list) else []:
        if not isinstance(item, dict):
            continue
        url = normalize_url(str(item.get("url", "")))
        if not url or url in seen:
            continue
        source_type = str(item.get("source_type", "other")).strip().lower()
        if source_type not in SOURCE_TYPES:
            source_type = "other"
        seen.add(url)
        sources.append(
            {
                "source_type": source_type,
                "publisher": str(item.get("publisher", "")).strip(),
                "title": str(item.get("title", "")).strip(),
                "published_date": str(item.get("published_date", "")).strip() or "Not disclosed",
                "url": url,
                "notes": str(item.get("notes", "")).strip(),
            }
        )
    return {
        "legal_name": str(payload.get("legal_name", "")).strip(),
        "official_website": official_website,
        "identity_confidence": str(payload.get("identity_confidence", "")).strip().lower(),
        "identity_notes": str(payload.get("identity_notes", "")).strip(),
        "sources": sources[:30],
    }


def clean_credit_reports(payload: Dict[str, Any]) -> List[Dict[str, str]]:
    reports: List[Dict[str, str]] = []
    seen = set()
    items = payload.get("reports", [])
    if not isinstance(items, list):
        return reports
    for item in items:
        if not isinstance(item, dict):
            continue
        url = normalize_url(str(item.get("report_url", "")))
        if not url or not _host_matches(url, RATING_AGENCY_DOMAINS) or url in seen:
            continue
        agency = str(item.get("agency", "")).strip()
        if not agency:
            continue
        seen.add(url)
        reports.append(
            {
                "agency": agency,
                "report_date": str(item.get("report_date", "")).strip() or "Not disclosed",
                "rating_action": str(item.get("rating_action", "")).strip() or "Not disclosed",
                "rating": str(item.get("rating", "")).strip() or "Not disclosed",
                "outlook": str(item.get("outlook", "")).strip() or "Not disclosed",
                "instrument": str(item.get("instrument", "")).strip() or "Not disclosed",
                "amount": str(item.get("amount", "")).strip() or "Not disclosed",
                "rationale": str(item.get("rationale", "")).strip(),
                "key_strengths": str(item.get("key_strengths", "")).strip(),
                "key_risks": str(item.get("key_risks", "")).strip(),
                "report_url": url,
            }
        )
    return reports[:20]


class ResearchOrchestrator:
    """Coordinate narrow source-discovery agents through the Responses API."""

    def __init__(self, api_key: str, model: str = DEFAULT_RESEARCH_MODEL):
        self.client = OpenAI(api_key=api_key, timeout=120.0, max_retries=2)
        self.model = model

    def _web_json(self, prompt: str, max_output_tokens: int = 3000) -> Dict[str, Any]:
        request: Dict[str, Any] = {
            "model": self.model,
            "input": prompt,
            "tools": [{"type": "web_search"}],
            "tool_choice": "auto",
            "max_tool_calls": 12,
            "max_output_tokens": max_output_tokens,
        }
        try:
            response = self.client.responses.create(**request)
        except TypeError:
            # Compatibility for older SDK versions. JSON is enforced in the
            # prompt and parsed defensively below.
            request.pop("max_tool_calls", None)
            response = self.client.responses.create(**request)
        return _extract_json_object(response.output_text or "")

    def _plain_json(self, prompt: str, max_output_tokens: int = 3000) -> Dict[str, Any]:
        request: Dict[str, Any] = {
            "model": self.model,
            "input": prompt,
            "max_output_tokens": max_output_tokens,
            "text": {"format": {"type": "json_object"}},
        }
        try:
            response = self.client.responses.create(**request)
        except TypeError:
            request.pop("text", None)
            response = self.client.responses.create(**request)
        return _extract_json_object(response.output_text or "")

    def discover_sources(self, company_name: str, cin: str = "") -> Dict[str, Any]:
        payload = self._web_json(
            f"""
You are the source-discovery agent for a transaction-grade company profile.

Company selected by the user: {company_name}
CIN, if known: {cin or "Not supplied"}

First disambiguate the exact legal entity. Find its genuine official website,
then locate high-value primary and reputable secondary sources: official pages,
annual reports, investor presentations, regulatory filings, company-registry
records, credit-rating reports, and reputable business news.

Return ONLY JSON:
{{
  "legal_name": "string",
  "official_website": "https://... or empty",
  "identity_confidence": "high|medium|low",
  "identity_notes": "brief disambiguation note",
  "sources": [
    {{
      "source_type": "official_website|credit_rating|annual_report|investor_presentation|regulatory_filing|registry|reputable_news|industry_source|other",
      "publisher": "string",
      "title": "string",
      "published_date": "YYYY-MM-DD, YYYY, or Not disclosed",
      "url": "https://...",
      "notes": "what this source can establish"
    }}
  ]
}}

Rules:
- Match the entity using legal name, CIN, address, directors, brands, and sector.
- Never treat a directory, marketplace, or social profile as the official site.
- Prefer first-party documents and direct document URLs.
- Include only URLs actually found in web search. Do not invent URLs.
- If identity is ambiguous, lower confidence and explain why.
""".strip(),
            max_output_tokens=4000,
        )
        return clean_source_dossier(payload)

    def discover_credit_ratings(
        self,
        company_name: str,
        cin: str = "",
        official_website: str = "",
    ) -> List[Dict[str, str]]:
        agencies = ", ".join(sorted(RATING_AGENCY_DOMAINS))
        payload = self._web_json(
            f"""
You are a credit-rating research specialist for Indian companies.

Company: {company_name}
CIN, if known: {cin or "Not supplied"}
Verified/likely official website: {official_website or "Not found yet"}

Search specifically for issuer rating rationales, rating action reports, press
releases, and surveillance/reaffirmation reports from official rating-agency
domains. Search current and historical names of the exact legal entity.

Approved rating-agency domains: {agencies}

Return ONLY JSON:
{{
  "reports": [
    {{
      "agency": "string",
      "report_date": "YYYY-MM-DD, YYYY, or Not disclosed",
      "rating_action": "assigned|upgraded|downgraded|reaffirmed|withdrawn|other",
      "rating": "rating symbols exactly as reported",
      "outlook": "string or Not disclosed",
      "instrument": "string or Not disclosed",
      "amount": "amount and currency or Not disclosed",
      "rationale": "concise, source-grounded rationale",
      "key_strengths": "semicolon-separated points",
      "key_risks": "semicolon-separated points",
      "report_url": "direct official agency URL"
    }}
  ]
}}

Rules:
- The report must refer to the exact company, not a similarly named entity.
- Use only approved official rating-agency domains.
- Do not infer a current rating from an old report; retain its report date.
- Preserve rating symbols, outlook, instrument, and amount exactly when found.
- Include only URLs actually found in web search. Do not invent URLs.
- Return an empty reports list when no exact match is found.
""".strip(),
            max_output_tokens=4500,
        )
        return clean_credit_reports(payload)

    def run_company_dossier(self, company_name: str, cin: str = "") -> Dict[str, Any]:
        dossier = self.discover_sources(company_name, cin)
        reports = self.discover_credit_ratings(
            company_name,
            cin=cin,
            official_website=dossier.get("official_website", ""),
        )
        return {"dossier": dossier, "credit_reports": reports}

    def build_mergermarket_search_plan(
        self,
        company_name: str,
        market_definition: Dict[str, Any],
    ) -> List[str]:
        """Create a compact set of Deals-screener queries without credentials."""
        payload = self._plain_json(
            f"""
You are designing searches for Mergermarket's Deals screener.

Subject company: {company_name}
Primary market: {market_definition.get("primary_market") or "Not established"}
Market description: {market_definition.get("market_description") or "Not established"}
Search keywords: {json.dumps(market_definition.get("keywords") or [], ensure_ascii=False)}
Excluded umbrella markets: {json.dumps(market_definition.get("exclude_markets") or [], ensure_ascii=False)}

Return ONLY JSON:
{{
  "queries": [
    "precise target-company description query for direct/core transactions",
    "broader product and customer-overlap query",
    "adjacent capability or value-chain query"
  ]
}}

Rules:
- Return 2 or 3 concise natural-language queries suitable for the Deals search box.
- Search target business descriptions, not the subject company name.
- Include distinctive products, technology, buyer use cases, or value-chain terms.
- Do not include dates or geography; the browser worker adds those deterministically.
- Do not include generic phrases such as industrial company, services company, or technology.
- Keep each query below 240 characters.
""".strip(),
            max_output_tokens=1200,
        )
        raw_queries = payload.get("queries", [])
        if not isinstance(raw_queries, list):
            return []
        queries: List[str] = []
        seen = set()
        for value in raw_queries:
            query = " ".join(str(value or "").split())[:240]
            key = query.lower()
            if len(query) >= 8 and key not in seen:
                seen.add(key)
                queries.append(query)
        return queries[:3]

    def classify_transactions(
        self,
        company_name: str,
        market_definition: Dict[str, Any],
        deals: List[Dict[str, Any]],
        batch_size: int = 25,
    ) -> List[Dict[str, Any]]:
        """Classify imported transactions without changing their source fields."""
        allowed_tiers = {"Core", "Broader", "Adjacent", "Excluded"}
        by_row = {str(deal.get("source_row")): deal for deal in deals}
        for start in range(0, len(deals), batch_size):
            batch = deals[start : start + batch_size]
            compact = [
                {
                    "row_ref": str(deal.get("source_row")),
                    "target": deal.get("target"),
                    "acquirer": deal.get("acquirer"),
                    "sector": deal.get("target_sector"),
                    "geography": deal.get("target_geography"),
                    "deal_type": deal.get("deal_type"),
                    "description": deal.get("description"),
                }
                for deal in batch
            ]
            payload = self._plain_json(
                f"""
You are the relevance-screening agent for precedent transaction comparables.

Subject company: {company_name}
Primary market: {market_definition.get("primary_market") or "Not established"}
Market description: {market_definition.get("market_description") or "Not established"}
Official website: {market_definition.get("official_website") or "Not established"}

Classify every candidate using its actual business description, not merely its
database sector label:
- Core: closely matches products, technology, customers, or value chain.
- Broader: same subsector with meaningful business-model overlap.
- Adjacent: informative context but less directly comparable.
- Excluded: not economically comparable, internal reorganisation, financing,
  unrelated asset, or otherwise unsuitable for valuation comparison.

CANDIDATES_JSON:
{json.dumps(compact, ensure_ascii=False)}

Return ONLY JSON:
{{
  "classifications": [
    {{
      "row_ref": "exact input row_ref",
      "relevance_tier": "Core|Broader|Adjacent|Excluded",
      "rationale": "specific one-sentence inclusion or exclusion rationale"
    }}
  ]
}}

Return exactly one classification per candidate. Do not add transactions or
alter deal facts.
""".strip(),
                max_output_tokens=max(1800, len(batch) * 110),
            )
            classifications = payload.get("classifications", [])
            if not isinstance(classifications, list):
                continue
            for item in classifications:
                if not isinstance(item, dict):
                    continue
                row_ref = str(item.get("row_ref", ""))
                tier = str(item.get("relevance_tier", "")).strip().title()
                if row_ref not in by_row or tier not in allowed_tiers:
                    continue
                by_row[row_ref]["relevance_tier"] = tier
                by_row[row_ref]["relevance_rationale"] = str(item.get("rationale", "")).strip()
        return deals
