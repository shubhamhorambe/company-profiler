# Company Intelligence Builder

A Streamlit application that combines PrivateCircle data with evidence-gated
web research to produce a sourced company intelligence workbook.

## Research coverage

- Legal-entity disambiguation and official company website discovery
- PrivateCircle financial statements
- Official credit-rating reports from supported Indian rating agencies
- Company products, customers, industry landscape, and market definition
- Competitors, M&A precedents, and prospective buyers
- Direct-source dossier for audit and follow-up research
- Claim-level evidence register across all researched sections
- Optional direct Mergermarket Deals screening with per-session credentials

The research agents only collect structured evidence. Workbook calculations,
source filtering, formatting, and export remain deterministic Python code.

## Local setup

Use Python 3.11, create a virtual environment, and install dependencies:

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

For local automatic Mergermarket retrieval, install Chrome/Chromium. If it is
not in a standard location, set `MERGERMARKET_BROWSER_EXECUTABLE` to its binary.
The deployed configuration installs Chromium through `packages.txt`.

Configure secrets in `.streamlit/secrets.toml`:

```toml
OPENAI_API_KEY = "..."
PC_API_KEY = "..."
APP_PASSWORD = "..." # strongly recommended for deployed instances
```

Then run:

```bash
streamlit run app.py
```

Environment variables with the same names are also supported. The optional
`PRIVATECIRCLE_API_KEY` name can be used instead of `PC_API_KEY`.

PrivateCircle statement values are normalised to INR millions using the unit
notes returned with each endpoint. For a legacy API account that returns raw
rupee values without reliable unit metadata, set `PC_FINANCIAL_INPUT_UNIT=inr`.
The default `auto` mode should be retained for current API responses.

## Mergermarket transaction comps

The app offers two transaction-data paths:

- **Public web research:** no Mergermarket access is required; the existing
  evidence-gated M&A workflow is used.
- **Connect Mergermarket automatically:** enter an authorised normal
  email/password login when the app opens. A fresh headless browser signs in,
  runs up to three agent-designed Deals searches, downloads the permitted
  Current Layout exports, parses them in memory, and closes. No manual Excel
  upload is required.

Credentials are never configured in `.env`, Streamlit secrets, source code,
logs, or generated workbooks. They exist only for the active Streamlit/browser
session. The worker does not bypass MFA, CAPTCHA, concurrent-session rules,
export allowances, or subscription controls. If Mergermarket cannot be used,
select Public web research or leave both login fields blank for that run.

The importer preserves the raw retrieved data, maps common column-name variants,
deduplicates by Deal ID and transaction identity, and uses a specialist agent
to classify Core, Broader, Adjacent, and Excluded candidates. The output adds
`Final Comps`, `Adjacent`, `Excluded`, `Sources`, `Search Log`, and `QA` sheets.
Calculated multiples remain formula-driven and are only produced from compatible
source fields. Public-web M&A research still runs as a complementary coverage
pass, so Mergermarket is a high-quality seed rather than the completeness limit.

The Mergermarket export dialog shows the account's remaining monthly allowance.
The app defaults to 100 deals per search and lets the user choose 50, 100, or
500. Keep the default unless the broader universe is genuinely needed.

## Model configuration

The defaults are set in `research_agents.py` and `pipeline_core_v2.py` and can
be changed without code edits:

```text
OPENAI_RESEARCH_MODEL=gpt-5.6-terra
OPENAI_REWRITE_MODEL=gpt-5.6-luna
```

Use a research model with Responses API web-search support. Before changing
models, run representative profiles and compare evidence quality, latency, and
cost.

## Validation

Run the offline test suite and syntax checks before deploying:

```bash
python -m unittest discover -s tests -v
python -m py_compile app.py pipeline_core_v2.py research_agents.py mergermarket_browser.py
```

Live end-to-end validation additionally requires valid OpenAI and PrivateCircle
credentials and consumes API usage.

## Security

Never commit API keys. Rotate any credential that has previously appeared in a
source file, shell history, shared workbook, or deployment log. For a production
multi-user deployment, replace the optional shared-password gate with your
organization's SSO or identity-aware proxy. Deploy this application privately;
each Mergermarket user must use an account whose licence permits automated
screening/export in this manner.
