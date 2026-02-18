# Input Schema

Use CSV headers:

`name,ir_url,ticker,country,cik`

Rules:

- Provide `name` or `ticker`.
- Provide `ir_url` whenever possible for reliable discovery in US and Canada.
- Provide `country` as `US` or `CA` to help source routing.
- Provide `cik` only for US issuers when known; otherwise ticker lookup is used.

Example CSV:

```csv
name,ir_url,ticker,country,cik
Agnico Eagle Mines,https://www.agnicoeagle.com/English/investors/default.aspx,AEM,CA,
Newmont Corporation,https://www.newmont.com/investors/default.aspx,NEM,US,1164727
Barrick Gold,https://www.barrick.com/English/investors/default.aspx,GOLD,CA,
```

Example command:

```bash
python scripts/fetch_latest_presentations.py --companies-csv companies.csv --output-dir output --latest-count 1
```

