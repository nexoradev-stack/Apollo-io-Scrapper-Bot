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
- An Apollo API key and the relevant endpoint/feature permissions for API-based Steps 1–2
- For the optional browser exporter: Playwright Chromium and an Apollo account you are authorised to use
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

3. Install the declared dependencies:

   ```bash
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

4. **Set the Apollo API key before running Step 1 or Step 2.** These scripts read the key from an environment variable and will stop if it is missing:

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

## Optional: browser-based visible-page export (no API key)

If you do not have an Apollo API key, you can try the visible-browser exporter. It opens a normal browser, lets you sign in manually, and collects only profile information rendered in the page. It does not intercept private API traffic or store your password in the script.

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   python -m playwright install chromium
   ```
2. Run:
   ```bash
   python scripts/STEP_1_apollo_browser_search.py
   ```
3. Paste an Apollo people-search URL, sign in in the opened browser if needed, and confirm the results are visible.
4. After each page, manually move to the next results page in the browser and press Enter in the terminal. Type `done` to export.

The exporter writes `output/apollo_browser_search_*.xlsx` with a Contacts sheet and a Summary sheet. Name, job title, company, and email fields depend on what Apollo visibly renders; title/company detection is heuristic, so review the workbook. If Apollo changes its page layout or blocks automated browsing, the script may need adjustment. This browser export is **not a replacement for API enrichment or CRM sync**; Steps 2 and other API-based features still require your own authorised API key. Use the browser exporter only in ways permitted by Apollo's terms and your account permissions.

## Workflow guide

Run the scripts in order, reviewing each workbook before proceeding:

1. **People search:** `python scripts/STEP_1_apollo_people_search.py`  
   Input: `inputs/companies_input.csv`  
   Output: `output/apollo_people_search_*.xlsx` and, if enabled, a master lead database.

2. **Title filtering (optional):** open the people-search workbook in Excel for the web. Follow `scripts/titles to be removed.md` and run `scripts/STEP_1point5_filter_search_results_by_titles.ts` from the Automate tab. Review results before enrichment.

3. **Enrichment / optional CRM sync:** by default, Step 2 automatically selects the latest `apollo_people_search_*.xlsx` in `output/`. To choose a specific file, set `APOLLO_ENRICH_INPUT_FILE` to its filename before running the script. Review the `Enriched`, `Matched`, `Unmatched`, and `CRM Save` sheets. CRM writes and personal-email reveal are disabled by default.

4. **Campaign format:** the default uses `output/master_lead_database.xlsx` (created by Step 2). Run `python scripts/STEP_3_apollo_to_ignite_format.py` to generate `output/ignite_ready_contacts_*.xlsx`. To use the latest enriched workbook instead, change `USE_MASTER_DB = False`; the newest enriched file is selected automatically, or set `APOLLO_ENRICHED_INPUT_FILE` to a filename.

5. **Master merge (optional):** place your own master workbook in `output/`, set `APOLLO_MERGE_MASTER_FILE` to its filename, and run `python scripts/STEP_4_apollo_ignite_master_merger.py`. Step 4 selects the newest `ignite_ready_contacts_*.xlsx` automatically; use `APOLLO_MERGE_SOURCE_FILE` to override it. Inspect the new merged workbook and keep a backup of the original master.

Detailed upstream usage notes: [How to Use Scrapper Utility](https://github.com/tejaspatel-jt/tj-apollo-data-scrapper/blob/master/How%20to%20Use%20Scrapper%20Utility.md).

## Configuration checklist

- Review input selection and workbook tab settings before each run. Steps 2–4 support auto-selecting the newest expected output when no override is configured.
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

README and workflow scripts have been added. Script configuration has been made safer, but live Apollo API runs, Microsoft Office Script execution, and end-to-end workbook processing still need testing in your environment.