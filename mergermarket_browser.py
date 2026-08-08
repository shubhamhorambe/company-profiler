"""Authenticated, ephemeral Mergermarket retrieval for authorised users.

The worker deliberately uses a fresh browser context for every run. Credentials
are accepted only as function arguments, are never written to disk, and are not
included in errors, logs, workbook metadata, or search records.
"""

from __future__ import annotations

import os
import re
import shutil
import time
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

from mergermarket_import import MergermarketImport, combine_mergermarket_imports, parse_mergermarket_export


MERGERMARKET_HOME = "https://mergermarket.ionanalytics.com/"
MERGERMARKET_DEALS = "https://mergermarket.ionanalytics.com/deals?version2"
DEALS_SEARCH_PLACEHOLDER = "Search Deals by Sector, Geography, Value, Multiples, Advisors..."


class MergermarketError(RuntimeError):
    """A safe-to-display Mergermarket workflow error."""


class MergermarketAuthenticationError(MergermarketError):
    pass


class MergermarketMFARequired(MergermarketError):
    pass


class MergermarketAccessError(MergermarketError):
    pass


class MergermarketRetrievalError(MergermarketError):
    pass


@dataclass(frozen=True)
class MergermarketCredentials:
    email: str
    password: str


@dataclass(frozen=True)
class MergermarketSearchPlan:
    queries: Sequence[str]
    start_date: date
    end_date: date
    geography_scope: str = "Domestic and global"
    max_deals_per_query: int = 100


@dataclass
class MergermarketBrowserResult:
    imported: MergermarketImport
    executed_queries: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


def redact_sensitive(value: object, secrets: Iterable[str]) -> str:
    """Remove credentials and common password representations from text."""
    text = str(value or "")
    for secret in secrets:
        if secret:
            text = text.replace(str(secret), "[REDACTED]")
    text = re.sub(r"(?i)(password|passwd|pwd)(\s*[:=]\s*)\S+", r"\1\2[REDACTED]", text)
    return text


def _browser_executable() -> Optional[str]:
    configured = os.getenv("MERGERMARKET_BROWSER_EXECUTABLE", "").strip()
    candidates = [
        configured,
        shutil.which("chromium") or "",
        shutil.which("chromium-browser") or "",
        shutil.which("google-chrome") or "",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ]
    return next((path for path in candidates if path and Path(path).exists()), None)


def _first_visible(locators) -> Optional[object]:
    for locator in locators:
        try:
            if locator.count() and locator.first.is_visible():
                return locator.first
        except Exception:
            continue
    return None


def _body_mentions(page, patterns: Sequence[str]) -> bool:
    try:
        body = page.locator("body").inner_text(timeout=2500).lower()
    except Exception:
        return False
    return any(pattern.lower() in body for pattern in patterns)


def _deal_frame(page):
    deadline = time.time() + 75
    while time.time() < deadline:
        for frame in page.frames:
            try:
                if frame.get_by_role("grid", name="Deals", exact=True).count():
                    return frame
                if frame.get_by_role("textbox", name=DEALS_SEARCH_PLACEHOLDER, exact=True).count():
                    return frame
            except Exception:
                continue
        time.sleep(0.5)
    raise MergermarketAccessError(
        "Mergermarket opened, but the Deals screener was not available for this account. "
        "Confirm that the subscription includes Deals access."
    )


def _submit_login(page) -> bool:
    button = _first_visible(
        [
            page.get_by_role("button", name=re.compile(r"^(sign in|log in|login|continue|next)$", re.I)),
            page.locator("button[type='submit']"),
            page.locator("input[type='submit']"),
        ]
    )
    if button is None:
        return False
    button.click()
    return True


def _authenticate(page, credentials: MergermarketCredentials) -> None:
    page.goto(MERGERMARKET_HOME, wait_until="domcontentloaded", timeout=90_000)

    for _ in range(5):
        if "mergermarket.ionanalytics.com" in page.url and _body_mentions(
            page, ["future deals", "auctions", "companies", "news"]
        ):
            return
        if _body_mentions(page, ["verification code", "authenticator code", "multi-factor", "two-factor"]):
            raise MergermarketMFARequired(
                "This Mergermarket account requires an additional verification step. "
                "The automatic worker does not bypass MFA; use the public-web fallback for this run."
            )
        if _body_mentions(page, ["captcha", "verify you are human", "unusual traffic"]):
            raise MergermarketAuthenticationError(
                "Mergermarket requested a human verification challenge. "
                "The worker will not bypass it; use the public-web fallback for this run."
            )

        email_input = _first_visible(
            [
                page.locator("input[type='email']"),
                page.locator("input[autocomplete='username']"),
                page.get_by_label(re.compile(r"email|username", re.I)),
                page.get_by_placeholder(re.compile(r"email|username", re.I)),
            ]
        )
        password_input = _first_visible(
            [
                page.locator("input[type='password']"),
                page.locator("input[autocomplete='current-password']"),
                page.get_by_label(re.compile(r"password", re.I)),
            ]
        )
        if email_input is not None:
            email_input.fill(credentials.email)
        if password_input is not None:
            password_input.fill(credentials.password)
        if email_input is None and password_input is None:
            time.sleep(1)
            continue
        _submit_login(page)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=30_000)
        except Exception:
            pass
        time.sleep(1)

    if _body_mentions(page, ["incorrect", "invalid credentials", "wrong password", "try again"]):
        raise MergermarketAuthenticationError(
            "Mergermarket rejected the email or password. Re-enter the credentials and try again."
        )
    raise MergermarketAuthenticationError(
        "Mergermarket sign-in did not complete. Confirm the email/password and that this account uses the normal login flow."
    )


def _query_text(plan: MergermarketSearchPlan, query: str) -> str:
    dates = f"Announcement Date {plan.start_date:%d %b %Y} to {plan.end_date:%d %b %Y}"
    if plan.geography_scope == "Domestic only":
        geography = "Target Geography India"
    elif plan.geography_scope == "Global only":
        geography = "Target Geography outside India"
    else:
        geography = "Target Geography India and global"
    return f"{query}; {geography}; {dates}"


def _export_current_results(page, frame, max_deals: int) -> Tuple[str, bytes]:
    export_button = frame.get_by_role("button", name="Export", exact=True)
    export_button.wait_for(state="visible", timeout=60_000)
    export_button.click(force=True)

    dialog = frame.get_by_role("dialog").first
    dialog.wait_for(state="visible", timeout=30_000)
    size = 500 if max_deals >= 500 else 100 if max_deals >= 100 else 50
    choice = dialog.get_by_role("radio", name=re.compile(rf"^First {size}"))
    if choice.count():
        choice.check(force=True)

    export_actions = dialog.get_by_role("button", name="Export", exact=True)
    with page.expect_download(timeout=180_000) as download_info:
        export_actions.last.click(force=True)
    download = download_info.value
    path = download.path()
    if not path:
        raise MergermarketRetrievalError("Mergermarket completed the export but no download file was returned.")
    data = Path(path).read_bytes()
    filename = download.suggested_filename or "mergermarket-deals.xlsx"
    if not data:
        raise MergermarketRetrievalError("Mergermarket returned an empty Deals export.")
    return filename, data


def fetch_mergermarket_transactions(
    credentials: MergermarketCredentials,
    plan: MergermarketSearchPlan,
) -> MergermarketBrowserResult:
    """Sign in, execute bounded Deals searches, export, parse, and close."""
    email = (credentials.email or "").strip()
    password = credentials.password or ""
    secrets = (email, password)
    if not email or not password:
        raise MergermarketAuthenticationError("Enter both the Mergermarket email and password.")
    if plan.end_date < plan.start_date:
        raise MergermarketRetrievalError("The transaction end date must be on or after the start date.")
    queries = [str(query).strip() for query in plan.queries if str(query).strip()][:3]
    if not queries:
        raise MergermarketRetrievalError("No usable Mergermarket search queries were generated.")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError as error:
        raise MergermarketRetrievalError(
            "The deployment is missing the browser worker dependency. Install the updated requirements and redeploy."
        ) from None

    imports: List[MergermarketImport] = []
    executed: List[str] = []
    warnings: List[str] = []
    try:
        with sync_playwright() as playwright:
            launch_options = {
                "headless": True,
                "args": ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
            }
            executable = _browser_executable()
            if executable:
                launch_options["executable_path"] = executable
            browser = playwright.chromium.launch(**launch_options)
            context = browser.new_context(accept_downloads=True, locale="en-US")
            page = context.new_page()
            page.set_default_timeout(45_000)
            try:
                _authenticate(page, MergermarketCredentials(email=email, password=password))
                page.goto(MERGERMARKET_DEALS, wait_until="domcontentloaded", timeout=90_000)
                frame = _deal_frame(page)
                search_box = frame.get_by_role("textbox", name=DEALS_SEARCH_PLACEHOLDER, exact=True)
                for query in queries:
                    full_query = _query_text(plan, query)
                    search_box.fill(full_query)
                    search_box.press("Enter")
                    frame.get_by_role("grid", name="Deals", exact=True).wait_for(state="visible", timeout=90_000)
                    time.sleep(3)
                    try:
                        filename, data = _export_current_results(page, frame, plan.max_deals_per_query)
                        imported = parse_mergermarket_export(filename, data)
                    except (ValueError, MergermarketRetrievalError) as error:
                        warnings.append(f"A Mergermarket search did not yield a usable export: {redact_sensitive(error, secrets)}")
                        try:
                            cancel = frame.get_by_role("button", name="Cancel", exact=True)
                            if cancel.count() and cancel.first.is_visible():
                                cancel.first.click(force=True)
                        except Exception:
                            pass
                        continue
                    imports.append(imported)
                    executed.append(query)
            finally:
                context.close()
                browser.close()
    except MergermarketError:
        raise
    except Exception as error:
        safe_detail = redact_sensitive(error, secrets)
        raise MergermarketRetrievalError(
            f"Mergermarket retrieval did not complete ({safe_detail[:240]}). "
            "No credentials were stored; use public-web research for this run if the site is temporarily unavailable."
        ) from None
    finally:
        del password
        del email

    if not imports:
        raise MergermarketRetrievalError(
            "Mergermarket returned no parseable Deals exports for the generated searches. "
            "Try a broader transaction scope or use public-web research for this run."
        )
    combined = combine_mergermarket_imports(imports, source_name="Automatic Mergermarket Deals searches")
    combined.warnings.extend(warnings)
    return MergermarketBrowserResult(imported=combined, executed_queries=executed, warnings=warnings)
