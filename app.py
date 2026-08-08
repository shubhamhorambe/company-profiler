import hmac
import logging
import os
import re

import streamlit as st

from mergermarket_import import parse_mergermarket_export
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
    "Do you have an authorised Mergermarket export for this profile?",
    options=["Public web research", "Mergermarket export + public enrichment"],
    horizontal=True,
    help="The app never stores or receives your Mergermarket password.",
)
transaction_scope = st.selectbox(
    "Transaction coverage",
    options=["Domestic and global", "Domestic only", "Global only"],
    index=0,
)
mergermarket_upload = None
mergermarket_preview = None
if transaction_source.startswith("Mergermarket"):
    st.info(
        "Sign in to Mergermarket in your own browser, export the filtered Deals results with the "
        "richest available columns, and upload the authorised export here."
    )
    mergermarket_upload = st.file_uploader(
        "Mergermarket Deals export",
        type=["xlsx", "csv"],
        help="Include Deal ID, parties, announcement date, status, description, value, financials, and multiples where available.",
    )
    if mergermarket_upload:
        if mergermarket_upload.size > 25 * 1024 * 1024:
            st.error("The Mergermarket export exceeds the 25 MB upload limit.")
        else:
            try:
                mergermarket_preview = parse_mergermarket_export(
                    mergermarket_upload.name,
                    mergermarket_upload.getvalue(),
                )
                st.success(
                    f"Loaded {len(mergermarket_preview.deals)} unique transaction candidates "
                    f"from sheet '{mergermarket_preview.source_sheet}'."
                )
                for warning in mergermarket_preview.warnings:
                    st.warning(warning)
            except Exception as error:
                st.error(f"The export could not be interpreted: {error}")

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
    requires_valid_export = transaction_source.startswith("Mergermarket") and mergermarket_preview is None
    generate = st.button(
        "Generate sourced profile",
        type="primary",
        use_container_width=True,
        disabled=requires_valid_export,
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
                    mergermarket_export_name=mergermarket_upload.name if mergermarket_preview else "",
                    mergermarket_export_bytes=mergermarket_upload.getvalue() if mergermarket_preview else None,
                    transaction_scope=transaction_scope,
                )
            st.session_state.generated_profile = {
                "company_id": selected["id"],
                "company_name": selected["name"],
                "data": workbook,
            }
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
