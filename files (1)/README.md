# The Open Jobs Map — UK

An explorer's map of advertised UK jobs, for people who don't yet know what they're
looking for. Fields of work are sized by how many roles are advertised and coloured by
pay (or by whether they're growing); each opens to show the roles inside it, their
trends since 2017, what the work is like, and how people get in.

It runs on free, openly-licensed **ONS** data (Open Government Licence): online job
advert volumes by occupation ("Labour demand volumes by SOC 2020"), with pay from ASHE
and role descriptions from National Careers Service profiles.

## Files
- `index.html` — the app. Reads `./data.json`; has an embedded copy as a fallback so it
  still renders if opened directly.
- `data.json` — the data the app reads. Ships with **modelled sample data** so the app
  works out of the box; replaced with live ONS data by the pipeline.
- `build_data.py` — downloads the latest ONS release and rebuilds `data.json` (live data).
- `gen_sample_data.py` — regenerates the sample `data.json` from the model (no internet).

## Run it locally
A browser won't `fetch` a local file over `file://`, so serve the folder:

```bash
cd jobsmap
python3 -m http.server 8000
# open http://localhost:8000
```

(If you skip this and just double-click `index.html`, it falls back to the embedded
sample data, which is fine for a look.)

## Make it live (real ONS data)
```bash
python3 -m pip install requests pandas openpyxl
python3 build_data.py --inspect          # shows the labour-demand workbook's sheets + columns
#   confirm the CONFIG block at the top of build_data.py matches what you see
python3 build_data.py                     # downloads counts (~45 MB) + ASHE pay (~80 MB), writes data.json
```
Reload the page — the sample-data flag disappears and you're on live figures. Re-run
`build_data.py` monthly (the counts/trends update; see the schedule step below).

### Pay (per-role, per-region medians)
Pay comes from **ASHE Table 15** (median annual pay by region × four-digit SOC), joined to
each role by SOC code. In the app, a role shows the **national** median under "All of the
UK" and its **regional** median when a region is selected, falling back to national where
ASHE has no figure for that role-in-region.

- By default `build_data.py` downloads ASHE and reads **national** medians automatically.
- For full **per-region** medians, the robust route is a CSV (ASHE zip layouts shift year
  to year): export "ASHE occupation (4-digit SOC)" with dimensions Median / Annual pay –
  Gross / All / each region from **NOMIS** (nomisweb.co.uk) or the ONS filter tool, save as
  `ashe.csv` with columns for SOC code, region and median, then:
  ```bash
  python3 build_data.py --ashe-csv ashe.csv
  ```
- Inspect the ASHE workbook layout any time with `python3 build_data.py --inspect-ashe`.
- `--no-ashe` skips it and uses bundled field medians (what the sample ships with).

ASHE refreshes once a year (each autumn), so you only redo the pay step annually.

> The downloads and field-grouping are robust. The bits most likely to need a one-time
> tweak are the labour-demand sheet/column names (`--inspect`) and, if you want per-region
> pay from the zip rather than a CSV, the ASHE workbook selection (`--inspect-ashe`). If a
> run errors, inspect and adjust the small `CONFIG` block. If the labour-demand file lacks
> a region breakdown, "Where" volumes fall back to population weights (the script says so).

## Deploy (free)
1. Put this folder in a Git repo (GitHub).
2. Connect it to **Vercel**, **Netlify** or **Cloudflare Pages** — they serve static
   sites free at this scale and deploy on every push.
3. Add a scheduled job that runs `build_data.py` and commits the updated `data.json`
   (GitHub Actions cron, or the host's scheduler).
4. Optional: point a custom domain at it (~£8–15/year).

Total running cost: **£0** at this scale; data is free under the Open Government Licence.

## Honest limits (shown in the app's "about" panel too)
- It shows **advertised demand, not total jobs** — one vacancy can appear as several
  adverts, and a large share of hiring is never advertised, so senior/network-driven
  roles are under-counted.
- **Monthly, not live**, with a ~4–6 week lag; pay refreshes yearly with ASHE.
- It can't show contract type, entry-level suitability, or links to apply — those would
  need a job-board feed layered on top.
- Fields are **occupational groups**, a clean rollup of SOC occupations (not industries).

## Attribution (required by the OGL)
> Contains public sector information licensed under the Open Government Licence v3.0.
> Source: Office for National Statistics.
