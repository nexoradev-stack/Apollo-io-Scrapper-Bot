# Apollo.io Scrapper Bot

A Python workflow for Apollo.io people search, enrichment, optional CRM sync, campaign-format exports, and workbook deduplication.

> **Source attribution:** This project is based on the public repository [tejaspatel-jt/tj-apollo-data-scrapper](https://github.com/tejaspatel-jt/tj-apollo-data-scrapper). It lives in a separate repository and is not a GitHub fork. Review the upstream licence and preserve required attribution before redistribution.

## Features

- Search for people and existing contacts by company.
- Export timestamped Excel workbooks and per-company summaries.
- Filter job titles using the included Microsoft Office Script.
- Enrich selected contacts using Apollo.
- Optionally save enriched contacts to Apollo CRM.
- Convert data to a campaign-ready workbook with country-to-region mapping.
- Merge records into a master workbook with deduplication.

## Requirements

- Python 3.10 or newer
- An Apollo API key and the relevant endpoint/feature permissions
- API credits for enrichment or data-reveal features you enable
- Excel for the web / Microsoft 365 to run the optional title-filter Office Script

## Setup

1. Clone the repository:

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

3. Install dependencies. Check the Python imports in the scripts and install their packages (the workflow primarily uses `requests`, `pandas`, and `openpyxl`):

   ```bash
   python -m pip install --upgrade pip
   pip install requests pandas openpyxl
   ```

4. **Secure the Apollo API key before running anything.** Upstream versions of the scripts contain a hard-coded API key. Do not run or redistribute scripts with embedded credentials. Replace hard-coded values with an environment variable, for example:

   ```python
   API_KEY = os.environ.get("APOLLO_API_KEY")
   if not API_KEY:
       raise RuntimeError("Set APOLLO_API_KEY before running this script")
   ```

   Set the variable in your terminal:

   **macOS/Linux**
   ```bash
   export APOLLO_API_KEY="your_apollo_api_key"
   ```

   **Windows PowerShell**
   ```powershell
   $env:APOLLO_API_KEY="your_apollo_api_key"
   ```

5. Review `inputs/companies_input.csv` and keep its column headers intact. Replace the example company list with companies you are authorised to research.

## Workflow guide

Run the scripts in order, reviewing each workbook before proceeding:

1. **People search:** `python scripts/STEP_1_apollo_people_search.py`  
   Input: `inputs/companies_input.csv`  
   Output: `output/apollo_people_search_*.xlsx` and, if enabled, a master lead database.

2. **Title filtering (optional):** open the people-search workbook in Excel for the web. Follow `scripts/titles to be removed.md` and run `scripts/STEP_1point5_filter_search_results_by_titles.ts` from the Automate tab. Review results before enrichment.

3. **Enrichment / CRM sync:** update the input workbook setting in `scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py`, then run it. Review the `Enriched`, `Matched`, `Unmatched`, and `CRM Save` sheets. Keep CRM creation disabled until you intentionally want records written to your Apollo account.

4. **Campaign format:** configure the input workbook in `scripts/STEP_3_apollo_to_ignite_format.py`, then run it to generate `output/ignite_ready_contacts_*.xlsx`.

5. **Master merge (optional):** configure the master and source workbook in `scripts/STEP_4_apollo_ignite_master_merger.py`, then run it. Inspect the newly generated merged workbook; retain a backup of your existing master data.

Detailed upstream usage notes: [How to Use Scrapper Utility](https://github.com/tejaspatel-jt/tj-apollo-data-scrapper/blob/master/How%20to%20Use%20Scrapper%20Utility.md).

## Configuration checklist

- Update example or stale workbook filenames in each stage before running it.
- Review all settings near the top of each script.
- Keep CRM create/write options disabled until you have reviewed test results and explicitly approve a write.
- Do not turn on personal-email or phone reveals unless your plan allows them and they are necessary.
- Start with a small test list to understand API credit usage and resulting output.

## Data and security

- Never commit API keys, tokens, production contact lists, Excel exports, or terminal logs.
- Use Apollo within the terms, API limits, and entitlements of your account.
- Ensure you have a lawful basis and permission to collect, enrich, store, and contact individuals.
- Honour applicable privacy, direct-marketing, retention, and opt-out requirements.
- The presence of a script or configuration toggle does not guarantee the associated API feature is enabled for your plan.

## Troubleshooting

- **401/403:** validate your key and Apollo plan permissions.
- **429:** reduce request rate and follow Apollo's retry guidance.
- **Input file not found:** update the file setting in the relevant script to match the workbook in `output/`.
- **Module not found:** activate the virtual environment and install the packages listed above.
- **Unexpected results:** test with a small input and check the workbook summary and unmatched records.

## References

- [Upstream source project](https://github.com/tejaspatel-jt/tj-apollo-data-scrapper)
- [Apollo API documentation](https://docs.apollo.io/)

## Project status

README and usage guidance have been added. The scripts still require a security/configuration review before production use; all features have not been independently tested in this repository.