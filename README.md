# Apollo.io Scrapper Bot

A Python workflow for researching people at target companies using Apollo.io, enriching selected records, optionally saving contacts to Apollo CRM, formatting records for campaign workflows, and merging results into a master workbook.

> **Source and attribution:** This repository is based on the public project [tejaspatel-jt/tj-apollo-data-scrapper](https://github.com/tejaspatel-jt/tj-apollo-data-scrapper). It is maintained as a separate repository, not a fork. Review the source project's licence and retain required attribution before redistributing or using derived code.

## Features

- Search Apollo people and contact records for companies from a CSV.
- Export search results to timestamped Excel workbooks.
- Filter job titles using the provided Microsoft Office Script.
- Enrich people with Apollo data.
- Optionally create contacts in Apollo CRM.
- Convert enriched records to a campaign-ready workbook with country-to-region mapping.
- Merge campaign-ready records with a master workbook using deduplication.
- Keep input and generated output files in separate folders.

## Requirements

- Python 3.10+ (Python 3.12 recommended)
- An Apollo API key with access to the endpoints/features you intend to use
- API access/credits for enrichment and any optional personal-email or phone reveal features
- Internet access
- For the title-filter step: Excel on the web or a compatible Microsoft 365 environment supporting Office Scripts

## Quick start

1. Clone this repository:

   ```bash
   git clone https://github.com/nexoradev-stack/Apollo-io-Scrapper-Bot.git
   cd Apollo-io-Scrapper-Bot
   ```

2. Create and activate a virtual environment:

   **macOS/Linux**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

   **Windows PowerShell**
   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. Install dependencies:

   ```bash
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

4. Configure your API key as an environment variable:

   **macOS/Linux**
   ```bash
   export APOLLO_API_KEY="your_apollo_api_key"
   ```

   **Windows PowerShell**
   ```powershell
   $env:APOLLO_API_KEY="your_apollo_api_key"
   ```

   Never commit API keys or real contact exports. Keep secrets in local environment variables or a secret manager.

5. Add target organisations to `inputs/companies_input.csv`. Keep the column headers expected by the script and use organisations you are authorised to research.

## Workflow

Run the stages in order, checking the output before moving to the next stage:

1. **People search** — run `python scripts/STEP_1_apollo_people_search.py`. Review the generated `output/apollo_people_search_*.xlsx` and summary.
2. **Title filtering (optional)** — upload the search workbook to Excel for the web, create a `titles` tab as described in `scripts/titles to be removed.md`, then run `scripts/STEP_1point5_filter_search_results_by_titles.ts` as an Office Script. Review the filtered file before enrichment.
3. **Enrichment / optional CRM sync** — set the input workbook in `scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py`, review the configuration flags, then run `python scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py`. Keep CRM creation disabled until you've verified the output and explicitly intend to write contacts into your Apollo account.
4. **Campaign format** — configure the intended master/enriched workbook in `scripts/STEP_3_apollo_to_ignite_format.py`, then run `python scripts/STEP_3_apollo_to_ignite_format.py`.
5. **Master merge (optional)** — configure the master and campaign-ready workbook in `scripts/STEP_4_apollo_ignite_master_merger.py`, then run `python scripts/STEP_4_apollo_ignite_master_merger.py`. Inspect the new merged workbook; the script is designed to write a separate output file.

Generated Excel workbooks and logs are saved under `output/`. Do not commit them.

## Important configuration check

The original scripts may contain embedded API keys and machine-specific/example filenames. **Before running any script**, inspect the configuration blocks. Replace any embedded key with `os.environ.get("APOLLO_API_KEY")` and fail fast if it is missing. Update input filenames to match the files in your local `output/` folder. Do not enable CRM writes by default: review the CRM output first and explicitly opt in.

## Repository layout

```text
inputs/       Company CSV input
scripts/      Python workflow stages and Office Script title filter
output/       Generated workbooks and logs (local; not tracked)
README.md     Setup and workflow guide
```

## Safety, compliance, and data handling

- Use the official Apollo API within the permissions, rate limits, terms, and plan entitlements of your account.
- Enrich and export only data you have a lawful basis and permission to process.
- Personal email and phone reveals may consume additional credits and should remain off unless explicitly needed and authorised.
- Validate records and CRM payloads before bulk writes; test with a small sample first.
- Never store API keys, private tokens, production contact lists, or personal data in Git.
- Apply relevant privacy, direct-marketing, retention, and opt-out requirements in your jurisdiction.

## Troubleshooting

- **401/403 errors:** check the API key, endpoint access, and Apollo plan permissions.
- **429/rate limits:** reduce request volume and respect retry/backoff guidance.
- **Missing input workbook:** update the configured filename in the relevant stage script.
- **Missing Python module:** activate the virtual environment and reinstall `requirements.txt`.
- **Unexpected output:** test with a small input CSV and inspect the workbook's Summary/Unmatched sheets before continuing.

## Upstream reference

- Source repository: https://github.com/tejaspatel-jt/tj-apollo-data-scrapper
- Apollo API documentation: https://docs.apollo.io/

## Status

Documentation is provided here. Review the upstream code's secrets/configuration and licensing before treating the workflow as ready for production.