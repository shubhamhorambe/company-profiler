import hmac
import logging
import os
import re
from datetime import date

import streamlit as st

from mergermarket_browser import MergermarketError
from pipeline_core_v2 import generate_profile_excel_bytes, privatecircle_search_dropdown


st.set_page_config(
    page_title="Company Intelligence Builder",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

logger = logging.getLogger(__name__)


def get_secret(key: str, default=None):
    """Read Streamlit secrets safely, with an environment-variable fallback."""
    try:
        value = st.secrets.get(key)
    except Exception:
        value = None
    return value if value not in (None, "") else os.getenv(key, default)


def require_password_if_configured() -> None:
    """Apply a simple deployment gate when APP_PASSWORD is configured."""
    expected = get_secret("APP_PASSWORD")
    if not expected or st.session_state.get("authenticated"):
        return

    st.title("Company Intelligence Builder")
    st.caption("Enter the deployment password to continue.")
    with st.form("login_form"):
        supplied = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in", type="primary")
    if submitted:
        if hmac.compare_digest(str(supplied), str(expected)):
            st.session_state.authenticated = True
            st.rerun()
        st.error("Incorrect password.")
    st.stop()


def option_label(company: dict) -> str:
    parts = [company.get("name")]
    location = ", ".join(value for value in (company.get("city"), company.get("state")) if value)
    if location:
        parts.append(location)
    if company.get("cin"):
        parts.append(f"CIN: {company['cin']}")
    if company.get("dba_name"):
        parts.append(f"DBA: {company['dba_name']}")
    return " | ".join(part for part in parts if part)


def safe_filename(company_name: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", company_name).strip("._")
    return f"{stem or 'Company'}_Intelligence_Profile.xlsx"


require_password_if_configured()

openai_api_key = get_secret("OPENAI_API_KEY")
privatecircle_api_key = get_secret("PC_API_KEY") or get_secret("PRIVATECIRCLE_API_KEY")

st.title("Company Intelligence Builder")
st.caption(
    "Build a sourced company profile using PrivateCircle financials, official company sources, "
    "credit-rating reports, and evidence-gated web research."
)

with st.sidebar:
    st.header("Research coverage")
    st.markdown(
        """
- Legal-entity and official-website verification
- Financial statements from PrivateCircle
- Credit-rating reports and rating actions
- Products, customers, and industry landscape
- Competitors, M&A transactions, and buyer universe
- Source dossier with direct links
"""
    )
    st.divider()
    st.caption("Outputs are research aids. Verify material conclusions against the linked primary sources.")

missing = []
if not openai_api_key:
    missing.append("OPENAI_API_KEY")
if not privatecircle_api_key:
    missing.append("PC_API_KEY (or PRIVATECIRCLE_API_KEY)")
if missing:
    st.error(f"Server configuration is incomplete: {', '.join(missing)}.")
    st.stop()

st.subheader("Transaction data source")
transaction_source = st.radio(
    "Transaction research mode",
    options=["Connect Mergermarket automatically", "Public web research"],
    horizontal=True,
    help="Automatic mode signs in for this run, exports screened Deals results, then closes its isolated browser session.",
)
transaction_scope = st.selectbox(
    "Transaction coverage",
    options=["Domestic and global", "Domestic only", "Global only"],
    index=0,
)
mergermarket_email = ""
mergermarket_password = ""
transaction_end_date = date.today()
try:
    transaction_start_date = transaction_end_date.replace(year=transaction_end_date.year - 10)
except ValueError:
    transaction_start_date = transaction_end_date.replace(year=transaction_end_date.year - 10, day=28)
mergermarket_max_deals = 100
if transaction_source.startswith("Connect Mergermarket"):
    st.info(
        "Enter your authorised Mergermarket login for this run. The credentials are held only in "
        "the active Streamlit session and are not written to disk or the workbook. The browser copy "
        "is discarded when its isolated session closes. The worker does not bypass MFA, CAPTCHA, or subscription controls."
    )
    st.caption("Leave both fields blank to use the normal public-web transaction workflow without Mergermarket.")
    mm_left, mm_right = st.columns(2)
    with mm_left:
        mergermarket_email = st.text_input(
            "Mergermarket email",
            placeholder="name@company.com",
            autocomplete="username",
        )
    with mm_right:
        mergermarket_password = st.text_input(
            "Mergermarket password",
            type="password",
            autocomplete="current-password",
        )
    dates = st.date_input(
        "Announcement-date period",
        value=(transaction_start_date, transaction_end_date),
        max_value=date.today(),
        help="The search agent applies this period to the Mergermarket Deals screener.",
    )
    if isinstance(dates, (tuple, list)) and len(dates) == 2:
        transaction_start_date, transaction_end_date = dates
    mergermarket_max_deals = st.select_slider(
        "Maximum deals exported per search",
        options=[50, 100, 500],
        value=100,
        help="The agent runs up to three focused searches. Larger exports consume more of the account's Mergermarket export allowance.",
    )

st.subheader("1. Find the legal entity")
with st.form("company_search_form"):
    query = st.text_input(
        "Company name",
        value=st.session_state.get("search_query", ""),
        placeholder="For example: Karam Safety Private Limited",
        help="Use the legal name where possible. A CIN makes identity matching more reliable.",
    )
    search = st.form_submit_button("Search PrivateCircle", type="primary")

if search:
    query = query.strip()
    st.session_state.search_query = query
    st.session_state.pop("generated_profile", None)
    if len(query) < 2:
        st.warning("Enter at least two characters.")
        st.session_state.pc_results = []
    else:
        with st.spinner("Searching PrivateCircle…"):
            try:
                st.session_state.pc_results = privatecircle_search_dropdown(
                    query,
                    privatecircle_api_key,
                    limit=15,
                )
            except Exception:
                logger.exception("PrivateCircle company search failed")
                st.session_state.pc_results = []
                st.error("Company search failed temporarily. Please retry in a moment.")

results = st.session_state.get("pc_results", [])
if search and not results:
    st.info("No exact matches were returned. Try a shorter legal name or remove punctuation.")

if results:
    st.subheader("2. Confirm the company")
    selected = st.selectbox(
        "PrivateCircle match",
        options=results,
        format_func=option_label,
        key="selected_company",
    )

    left, right = st.columns(2)
    left.metric("Selected company", selected.get("name") or "—")
    right.metric("CIN", selected.get("cin") or "Not available")

    st.subheader("3. Generate the intelligence workbook")
    st.write(
        "Generation can take several minutes because the app verifies sources and runs multiple "
        "specialist research stages."
    )
    has_mm_email = bool(mergermarket_email.strip())
    has_mm_password = bool(mergermarket_password)
    incomplete_credentials = transaction_source.startswith("Connect Mergermarket") and has_mm_email != has_mm_password
    if incomplete_credentials:
        st.warning("Enter both Mergermarket fields, or leave both blank to use public-web research.")
    generate = st.button(
        "Generate sourced profile",
        type="primary",
        use_container_width=True,
        disabled=incomplete_credentials,
    )

    if generate:
        progress = st.progress(0)
        status = st.empty()

        def progress_callback(message, percent):
            status.write(message)
            progress.progress(max(0, min(100, int(percent))))

        try:
            with st.spinner("Running company research…"):
                workbook = generate_profile_excel_bytes(
                    company_name=selected["name"],
                    encrypted_id=selected["id"],
                    openai_api_key=openai_api_key,
                    privatecircle_token=privatecircle_api_key,
                    progress_cb=progress_callback,
                    cin=selected.get("cin") or "",
                    transaction_scope=transaction_scope,
                    mergermarket_email=mergermarket_email if has_mm_email and has_mm_password else "",
                    mergermarket_password=mergermarket_password if has_mm_email and has_mm_password else "",
                    transaction_start_date=transaction_start_date,
                    transaction_end_date=transaction_end_date,
                    mergermarket_max_deals_per_query=mergermarket_max_deals,
                )
            st.session_state.generated_profile = {
                "company_id": selected["id"],
                "company_name": selected["name"],
                "data": workbook,
            }
        except MergermarketError as error:
            logger.warning("Mergermarket workflow did not complete: %s", error)
            st.error(str(error))
            st.info("Switch Transaction research mode to Public web research to generate the profile without Mergermarket.")
        except Exception:
            logger.exception("Profile generation failed for company id %s", selected.get("id"))
            st.error(
                "Profile generation did not complete. No partial workbook was downloaded. "
                "Please retry; if the issue persists, check the deployment logs."
            )

    generated = st.session_state.get("generated_profile")
    if generated and generated.get("company_id") == selected.get("id"):
        st.success("Profile completed. The workbook includes source and credit-rating audit sheets.")
        st.download_button(
            label="Download Excel profile",
            data=generated["data"],
            file_name=safe_filename(generated["company_name"]),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
