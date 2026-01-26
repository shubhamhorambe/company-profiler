import streamlit as st

from pipeline_core_v2 import privatecircle_search_dropdown, generate_profile_excel_bytes

import os

def get_secret(key: str):
    try:
        return st.secrets[key]        # Streamlit Cloud or local secrets.toml
    except Exception:
        return os.getenv(key)         # Local / Codespaces fallback

OPENAI_API_KEY = get_secret("OPENAI_API_KEY")
PC_API_KEY = get_secret("PC_API_KEY")  # or PRIVATECIRCLE_API_KEY if that’s what you standardise on

if not OPENAI_API_KEY:
    st.error("OPENAI_API_KEY not set (Streamlit secrets or environment variable).")
    st.stop()


st.set_page_config(page_title="Company Profiler", layout="wide")
st.title("Company Profiler")


# If you’re using auth elsewhere, you can keep these; not used in this minimal file:
COOKIE_SECRET = st.secrets.get("COOKIE_SECRET")
APP_USERS_JSON = st.secrets.get("APP_USERS_JSON")

if not OPENAI_API_KEY or not PC_API_KEY:
    st.error("Server configuration error: OPENAI_API_KEY / PC_API_KEY not set in Streamlit Secrets.")
    st.stop()

st.subheader("1) Search company on PrivateCircle")

query = st.text_input(
    "Type company name and press Enter",
    placeholder="e.g., Karam Safety",
)

if "pc_results" not in st.session_state:
    st.session_state.pc_results = []

# Search whenever query changes
if query:
    with st.spinner("Searching PrivateCircle..."):
        try:
            st.session_state.pc_results = privatecircle_search_dropdown(query, PC_API_KEY, limit=15)
        except Exception as e:
            st.session_state.pc_results = []
            st.error(f"PrivateCircle search failed: {e}")
else:
    st.session_state.pc_results = []

results = st.session_state.pc_results


def option_label(x):
    bits = [x.get("name")]
    loc = ""
    if x.get("city") and x.get("state"):
        loc = f"{x.get('city')}, {x.get('state')}"
    elif x.get("city"):
        loc = x.get("city")
    elif x.get("state"):
        loc = x.get("state")
    if loc:
        bits.append(loc)
    if x.get("cin"):
        bits.append(f"CIN: {x['cin']}")
    if x.get("dba_name"):
        bits.append(f"DBA: {x.get('dba_name')}")
    return " | ".join([b for b in bits if b])


if query and not results:
    st.warning("No matches found. Try another spelling, shorter keyword, or partial name.")
elif results:
    st.subheader("2) Select the correct company")

    options = [option_label(x) for x in results]
    choice = st.selectbox("Select company match", options, index=0)
    selected = results[options.index(choice)]

    st.success(f"Selected: **{selected['name']}**")
    encrypted_id = selected["id"]
    company_name = selected["name"]

    st.subheader("3) Generate profile")

    generate = st.button("Generate Profile", type="primary", use_container_width=True)

    if generate:
        prog = st.progress(0)
        status = st.empty()

        def progress_cb(msg, pct):
            status.write(msg)
            try:
                prog.progress(int(pct))
            except Exception:
                pass

        with st.spinner("Running pipeline..."):
            excel_bytes = generate_profile_excel_bytes(
                company_name=company_name,
                encrypted_id=encrypted_id,
                openai_api_key=OPENAI_API_KEY,
                privatecircle_token=PC_API_KEY,
                progress_cb=progress_cb,
            )

        st.success("Completed. Download your Excel below.")
        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name=f"{company_name}_Profile.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
