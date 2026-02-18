---
name: mining-presentation-fetcher
description: Download the latest investor presentation files and full extracted text for publicly traded mining companies in the US and Canada. Use when asked to fetch the newest company deck(s), archive latest presentation content, or capture complete text from recent investor presentations without doing valuation or analysis.
---

# Mining Presentation Fetcher

## Overview

Use `scripts/fetch_latest_presentations.py` to find latest investor presentation files, download them, and extract full text into local files.
Prioritize company IR pages for all markets and optionally query SEC filings for US-listed names.

## Workflow

1. Prepare company inputs as either shorthand `ticker|exchange` or full fields with `name`, `ir_url`, `ticker`, `country`, `cik`, and `exchange`.
2. Run `scripts/fetch_latest_presentations.py` with `--companies-csv` or repeated `--company` flags.
3. Crawl IR pages, score candidate presentation links, and pick the most recent files by date signals.
4. Download latest file(s) and extract full text from `pdf` and `pptx` documents.
5. Save outputs and manifest files under the chosen output directory.

## Run Commands

Install dependencies:

```bash
pip install requests beautifulsoup4 pypdf python-pptx
```

Optional browser fallback for IR sites that block direct HTTP crawling:

```bash
pip install playwright
playwright install chromium
```

Fetch one latest presentation per company from CSV:

```bash
python scripts/fetch_latest_presentations.py --companies-csv companies.csv --output-dir output --latest-count 1
```

Fetch from inline company specs and include SEC fallback for US tickers:

```bash
python scripts/fetch_latest_presentations.py \
  --company "AEM|TSX" \
  --company "NEM|NYSE" \
  --include-sec \
  --playwright-fallback \
  --output-dir output
```

Preview discovered links without downloading:

```bash
python scripts/fetch_latest_presentations.py --companies-csv companies.csv --dry-run --verbose
```

## Web App

Install web dependency:

```bash
pip install -r web/requirements.txt
```

Run local UI:

```bash
python web/app.py
```

Open:

```text
http://127.0.0.1:5050
```

## Output Contract

- Save downloaded files under `output/<company-slug>/`.
- Save extracted text as sibling `*.txt` files.
- Write aggregate manifests:
- `output/manifest.json`
- `output/manifest.csv`

## Notes

- Keep focus on collection and text capture only.
- Avoid valuation, earnings interpretation, or investment recommendations in this skill.
- Read `references/input_schema.md` for CSV format and usage constraints.
