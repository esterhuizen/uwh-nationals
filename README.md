# UWH Nationals 2026 — live results

Live scores and draw for the **NZ Secondary Schools Underwater Hockey Nationals 2026**
(Rotorua Aquatic Centre, 3–6 September 2026), with the two tracked teams —
**JOB HC-BD** and **SGB Howick** — pulled out at the top.

**Page: https://esterhuizen.github.io/uwh-nationals/**

## How it updates

A GitHub Action (`.github/workflows/update.yml`) runs every 20 minutes and on demand
(`workflow_dispatch`). It executes `scripts/fetch_data.py`, which:

1. Downloads the public tournament workbook from Google Sheets
   (3 attempts, 5s/15s backoff, verified as a real `.xlsx`).
2. Parses the `Inputs` sheet — the master draw, two courts side by side — into every game:
   day, date, time (NZST), court, White/Black teams, scores, stage and division.
3. Resolves knockout placeholders such as `SG Winner G85` to the actual team once the
   referenced game has been played and was not a draw (the original label is kept in
   `white_source` / `black_source`).
4. Writes `docs/data.json` only when something actually changed, ignoring the
   `generated_at` timestamp — so quiet periods produce no commits.

If the download fails, or the parse yields implausibly few games, the script exits
non-zero and leaves the last good `docs/data.json` untouched.

The site itself is `docs/index.html` — one self-contained file, no frameworks, no CDNs.
It fetches `data.json` on load and re-fetches every 5 minutes.

## Running it locally

```sh
pip install openpyxl
python scripts/fetch_data.py                 # from the live Google Sheet
python scripts/fetch_data.py --source draw.xlsx   # from a local copy
```

Prints `CHANGED` or `UNCHANGED`.

## Layout

```
scripts/fetch_data.py        fetch + parse + write docs/data.json
docs/index.html              the site
docs/data.json               generated data (committed by the Action)
docs/.nojekyll               serve files starting with _ / skip Jekyll
.github/workflows/update.yml the 20-minute refresh
```

Published with GitHub Pages from `main` → `/docs`.
