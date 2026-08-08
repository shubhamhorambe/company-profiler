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
- Optional authorised Mergermarket export ingestion and relevance screening

The research agents only collect structured evidence. Workbook calculations,
source filtering, formatting, and export remain deterministic Python code.

## Local setup

Use Python 3.11, create a virtual environment, and install dependencies:

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

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
- **Mergermarket export + public enrichment:** sign in to Mergermarket in your
  own browser, export authorised Deals results as `.xlsx` or `.csv`, and upload
  the file. The app does not receive or store your Mergermarket password.

The importer preserves the raw export, maps common column-name variants,
deduplicates by Deal ID and transaction identity, and uses a specialist agent
to classify Core, Broader, Adjacent, and Excluded candidates. The output adds
`Final Comps`, `Adjacent`, `Excluded`, `Sources`, `Search Log`, and `QA` sheets.
Calculated multiples remain formula-driven and are only produced from compatible
source fields.

An official Mergermarket API or service-account integration can be added behind
the same import boundary when licensed API documentation is available.

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
python -m py_compile app.py pipeline_core_v2.py research_agents.py
```

Live end-to-end validation additionally requires valid OpenAI and PrivateCircle
credentials and consumes API usage.

## Security

Never commit API keys. Rotate any credential that has previously appeared in a
source file, shell history, shared workbook, or deployment log. For a production
multi-user deployment, replace the optional shared-password gate with your
organization's SSO or identity-aware proxy.
