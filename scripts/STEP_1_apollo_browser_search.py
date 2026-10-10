#!/usr/bin/env python3
"""Visible-page Apollo search export using Playwright.

This script does not use an Apollo API key, inspect private API traffic, or
collect login credentials. The user signs in manually in a visible browser.
Only information rendered in the page DOM is exported.
"""

from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

ROOT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT_DIR / "output"
PROFILE_DIR = Path.home() / ".apollo_scrapper_browser_profile"


def extract_visible_people(page) -> list[dict[str, str]]:
    """Extract visible person cards anchored by Apollo profile links."""
    return page.evaluate(
        r"""() => {
          const visible = (el) => {
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 &&
              s.visibility !== 'hidden' && s.display !== 'none';
          };
          const links = [...document.querySelectorAll('a[href]')]
            .filter(a => visible(a) && /\/people\//i.test(a.getAttribute('href') || ''));
          const seen = new Set();
          const rows = [];
          for (const a of links) {
            const href = new URL(a.getAttribute('href'), location.href).href;
            if (seen.has(href)) continue;
            seen.add(href);
            let node = a;
            let best = a;
            for (let i = 0; i < 7 && node.parentElement; i++) {
              node = node.parentElement;
              const text = (node.innerText || '').trim();
              if (text.length >= 30 && text.length <= 1800 &&
                  text.split('\n').filter(x => x.trim()).length >= 3) {
                best = node;
                break;
              }
            }
            const raw = (best.innerText || a.innerText || '').trim();
            if (!raw) continue;
            const lines = raw.split('\n').map(x => x.trim()).filter(Boolean);
            const emails = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig) || [];
            const linkedin = [...best.querySelectorAll('a[href]')]
              .map(x => x.href).find(x => /linkedin\.com\/in\//i.test(x)) || '';
            rows.push({
              'Name': (a.innerText || '').trim(),
              'Job Title': '',
              'Company': '',
              'Email': [...new Set(emails)].join('; '),
              'LinkedIn URL': linkedin,
              'Apollo Profile URL': href,
              'Visible Card Text': raw
            });
            const row = rows[rows.length - 1];
            const name = row['Name'];
            const nameIndex = name ? lines.findIndex(x => x === name) : -1;
            const remaining = nameIndex >= 0 ? lines.filter((_, i) => i !== nameIndex) : lines;
            // Conservative heuristics only; keep the full visible card text for review.
            const titleLine = remaining.find(x =>
              /\b(chief|ceo|cto|cfo|cmo|cio|vp|vice president|director|head|manager|founder|president|engineer|marketing|sales|product)\b/i.test(x)
              && x.length < 180
            );
            if (titleLine) row['Job Title'] = titleLine;
            const companyLine = remaining.find(x =>
              x.length < 140 && x !== titleLine &&
              !/@/.test(x) && !/^(view|connect|message|save|more|linkedin)$/i.test(x)
            );
            if (companyLine) row['Company'] = companyLine;
          }
          return rows;
        }"""
    )


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    search_url = input(
        "Paste your Apollo people-search URL (you can sign in when the browser opens): "
    ).strip()
    if not search_url.startswith("https://"):
        print("Please provide an HTTPS Apollo search URL.")
        return 2
    host = (urlparse(search_url).hostname or "").lower()
    if not (host == "apollo.io" or host.endswith(".apollo.io")):
        print("For safety, the URL must be on apollo.io.")
        return 2

    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    all_rows: dict[str, dict[str, str]] = {}
    timestamp = datetime.now().strftime("%d%b%Y_%H%M%S").lower()
    output_file = OUTPUT_DIR / f"apollo_browser_search_{timestamp}.xlsx"

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1440, "height": 1000},
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(search_url, wait_until="domcontentloaded")
        print("\n1. Sign in to Apollo manually if prompted.")
        print("2. Apply/confirm your filters and wait for the results to load.")
        input("When the people-search results are visible, press Enter here... ")

        while True:
            try:
                page.wait_for_load_state("domcontentloaded", timeout=5000)
            except PlaywrightTimeoutError:
                pass
            rows = extract_visible_people(page)
            for row in rows:
                key = row.get("Apollo Profile URL") or (
                    row.get("Name", "") + "|" + row.get("Visible Card Text", "")
                )
                if key.strip("|"):
                    all_rows[key] = row
            print(f"Captured {len(rows)} visible profile cards on this page; "
                  f"{len(all_rows)} unique profiles collected so far.")
            action = input(
                "Go to the next Apollo results page in the browser, then press Enter; "
                "type 'done' to export: "
            ).strip().lower()
            if action == "done":
                break

        context.close()

    if not all_rows:
        print(
            "\nNo person-profile cards were detected in the visible page. "
            "Apollo may have changed its page layout; no file was created."
        )
        return 1

    df = pd.DataFrame(all_rows.values())
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Contacts", index=False)
        pd.DataFrame(
            [
                {"Metric": "Unique visible profiles", "Value": len(df)},
                {"Metric": "Exported at", "Value": datetime.now().isoformat(timespec="seconds")},
                {"Metric": "Collection method", "Value": "Visible browser DOM; manual pagination"},
            ]
        ).to_excel(writer, sheet_name="Summary", index=False)

    print(f"\nExported {len(df)} visible profiles to: {output_file}")
    print("Review the workbook: title/company fields are heuristic and may need correction.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
