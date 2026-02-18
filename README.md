# Mining Presentation Fetcher

Fetch the latest investor presentation files for public mining companies and save full extracted text.

The project supports:
- Web UI workflow (ticker + exchange form)
- CLI workflow (single or multiple companies)
- US SEC fallback for US-listed tickers
- Optional browser automation fallback for blocked investor-relations pages

## 1) What This Project Does

For each company, the fetcher:
1. Discovers candidate presentation files (`.pdf`, `.ppt`, `.pptx`) from IR pages and search fallbacks.
2. Ranks candidates by date and relevance keywords.
3. Selects the latest `N` files.
4. Downloads files.
5. Extracts full text to `.txt`.
6. Writes metadata and manifest files.

This tool is for collection and text capture, not investment analysis.

## 2) Main Workflow (Web UI)

### Step A: Install dependencies

```bash
pip install -r web/requirements.txt
pip install requests beautifulsoup4 pypdf python-pptx
```

Optional (recommended for sites that block normal HTTP requests):

```bash
pip install playwright
playwright install chromium
```

### Step B: Start the app

```bash
python web/app.py
```

Open:

`http://127.0.0.1:5050`

### Step C: Fill the form

Primary input:
- `Ticker` (example: `AEM`)
- `Exchange` (select from dropdown)

Supported exchanges in UI:
- `NYSE`
- `NASDAQ`
- `AMEX`
- `TSX`
- `TSXV`
- `CSE`
- `ASX`
- `LSE`

Optional input:
- Upload CSV for batch runs (`name,ir_url,ticker,country,cik,exchange`)

Run options:
- `Latest count`: how many latest files per company
- `Timeout (sec)`: request timeout
- `Include SEC fallback`: query SEC submissions for US companies
- `Playwright fallback`: use browser rendering for blocked pages
- `Disable web search fallback`: disable DDG/Bing fallback
- `Dry run`: discover only, no downloads

### Step D: Run and review output

After submitting:
- `Run Result` shows exit code and record count
- `Files` table shows links to deck and text files
- `Logs` shows stdout/stderr from the fetch process

## 3) CLI Workflow

### 3.1 Single company (short format)

```bash
python scripts/fetch_latest_presentations.py --company "AEM|TSX" --output-dir output --latest-count 1
```

### 3.2 Single company (full format)

```bash
python scripts/fetch_latest_presentations.py --company "Agnico Eagle Mines|https://www.agnicoeagle.com/English/investors/default.aspx|AEM|CA||TSX" --output-dir output --latest-count 1
```

Full format:

`name|ir_url|ticker|country|cik|exchange`

### 3.3 Batch companies from CSV

```bash
python scripts/fetch_latest_presentations.py --companies-csv companies.csv --output-dir output --latest-count 1
```

CSV headers:

`name,ir_url,ticker,country,cik,exchange`

### 3.4 Useful flags

```bash
python scripts/fetch_latest_presentations.py \
  --company "NEM|NYSE" \
  --latest-count 2 \
  --include-sec \
  --playwright-fallback \
  --timeout 20 \
  --verbose \
  --output-dir output
```

Flag summary:
- `--company`: repeat for multiple companies
- `--companies-csv`: batch input file
- `--latest-count`: latest files per company
- `--include-sec`: SEC fallback for US names
- `--playwright-fallback`: browser fallback on blocked pages
- `--disable-search-fallback`: disable DDG/Bing fallback
- `--dry-run`: discover without download
- `--timeout`: network timeout in seconds
- `--verbose`: debug logging

## 4) Data Discovery and Selection Logic

Per company, discovery runs in this order:
1. IR crawl (if IR URL exists):
   - Seed page and common IR path variants
   - Links scored by presentation keywords and file extension
2. SEC fallback (if enabled and company is US-listed)
3. Web search fallback (unless disabled)

Candidate ranking:
- Primary: parsed date (newer first)
- Secondary: keyword score

Keywords include terms like `presentation`, `investor`, `deck`, `corporate`, `factsheet`.

## 5) Output Structure

### CLI default output

`output/`
- `<company-slug>/`
- `<date>__<title>.pdf|ppt|pptx`
- `<date>__<title>.txt`
- `<date>__<title>.metadata.json`
- `manifest.json`
- `manifest.csv`

### Web UI run output

`web/runs/<job_id>/`
- `run.json` (command, options, logs, return code)
- `output/manifest.json`
- `output/manifest.csv`
- company folders and files

Metadata fields include:
- `company_name`
- `ticker`
- `exchange`
- `country`
- `source`
- `source_page`
- `document_url`
- `published_date`
- `local_file`
- `text_file`
- `text_extraction_method`
- `text_char_count`

## 6) Running in Background (Windows)

Start app in background:

```powershell
$proc = Start-Process -FilePath python -ArgumentList 'web/app.py' -WorkingDirectory 'c:\dev\Miner_skills\mining-presentation-fetcher' -PassThru
$proc.Id
```

Check app health:

```powershell
(Invoke-WebRequest -Uri 'http://127.0.0.1:5050' -UseBasicParsing).StatusCode
```

Stop app:

```powershell
Stop-Process -Id <PID>
```

## 7) Troubleshooting

### Problem: No presentation candidates found

Try:
1. Enable `Playwright fallback`
2. Increase `Timeout`
3. Keep `web search fallback` enabled
4. Provide a known IR URL via full format or CSV

### Problem: SEC fallback returns nothing

Checks:
1. Exchange/country is US (`NYSE`, `NASDAQ`, `AMEX`, or `country=US`)
2. Ticker is correct
3. Company has recent 8-K/6-K records containing presentation attachments

### Problem: PDF text extraction is weak

Notes:
- Scanned-image PDFs may return low text content with `pypdf`.
- OCR is not included in current version.

## 8) Development Notes

Core files:
- `scripts/fetch_latest_presentations.py`: main fetch engine
- `web/app.py`: Flask web app
- `web/templates/index.html`: UI markup
- `web/static/styles.css`: UI styles
- `references/input_schema.md`: input schema reference

Validation:

```bash
python -m py_compile scripts/fetch_latest_presentations.py web/app.py
python C:/Users/User/.codex/skills/.system/skill-creator/scripts/quick_validate.py c:/dev/Miner_skills/mining-presentation-fetcher
```

