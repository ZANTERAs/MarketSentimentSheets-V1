# MarketSentimentSheets-V1
# HARM News Bot (Excel Portfolio)

Collect, tag, score, and export market news to Excel—ready for portfolio tracking and analysis.

This repo contains a small toolkit centered on **`news_harm.py`** (the bot) and **`ticker_aliases.py`** (helper to build name aliases). A static **`aliases.json`** is also included.

---

## Files in this folder

- **`news_harm.py`** – Main script: fetches RSS news, tags companies/tickers, scores sentiment, and writes Excel workbooks with formatting. fileciteturn3file1
- **`ticker_aliases.py`** – CLI helper to build/update a JSON of ticker → name/alias variants using Yahoo Finance. fileciteturn3file2
- **`aliases.json`** – Example mapping of tickers to readable names/aliases used for matching in headlines/summaries. fileciteturn3file0

Output directory created on first run: **`news_bot_output/`**.

---

## Key features

- **Company tagging**: finds tickers in headlines/summaries using aliases (e.g., “Exelon”, “EXC”).  
- **Sentiment scoring**: VADER by default; FinBERT optional.  
- **Three Excel outputs** (overwritable or persistent):
  1. `news_db.xlsx` – persistent database (deduped by unique ID). Columns: **Date, Source, Title, Link, Sentiment, ID, Tickers, Companies**.
  2. `news_last7.xlsx` – rolling **last _N_ days** view with summaries. Columns: **Date, Source, Title, Summary, Link, Sentiment, ID, Tickers, Companies**.
  3. `news_outputs.xlsx` – legacy multi‑sheet (**RawNews**, **MappedScored**, **DailySignals**).
- **Excel formatting** (applied automatically):
  - `Summary` ≈ **500 px**, `Source` ≈ **220 px**, `Date` ≈ **80 px**
  - Dates **render as `YYYY‑MM‑DD`** (no time)
  - **Thin external borders** around each data row
- **Safe, growing DB**: appends on every run and de‑duplicates by **ID** (MD5 of title+summary+link).

---

## Requirements

- **Python**: 3.10–3.12 recommended
- **Packages**:
  ```bash
  pip install pandas openpyxl feedparser nltk yfinance
  # optional (plots & alt sentiment)
  pip install plotly transformers torch
  ```
- First VADER run will download the lexicon automatically.

> Note: Some libraries may lag Python 3.13; use 3.12 if you hit install/runtime issues.

---

## Quick start

From the folder containing `news_harm.py`:

```powershell
# Windows PowerShell
py .\news_harm.py --tickers EXC XEL MSFT --backend vader --days 7
```

Outputs are written to **`news_bot_output/`**:
- `news_db.xlsx` (persistent DB)
- `news_last7.xlsx` (last N days + summary)
- `news_outputs.xlsx` (legacy multi‑sheet)

> **Tip**: If you see “No module named news_harm”, run it as a script from this folder (use `py .\news_harm.py ...`), not `-m`.

---

## CLI: `news_harm.py`

```text
--tickers   List of tickers to track (e.g., EXC XEL MSFT)
--backend   vader | finbert  (default: vader)
--days      Rolling window for "last N days" view (default: 7)
--plot      If set, shows sentiment vs returns (optional)
--lookahead Lookahead days for forward returns in the legacy sheet (default: 1)
```

**Important:** `--days` filters what you *show*, not what RSS returns. Most RSS feeds only expose the latest ~20–100 items. For true last‑365 coverage, either let the **DB accumulate** across runs and build the view from the DB, or integrate a historical news API.

---

## CLI: `ticker_aliases.py`

Builds a JSON mapping _ticker → alias list_ via Yahoo Finance metadata.

```bash
# Default (uses built-in DEFAULT_TICKERS)
python ticker_aliases.py --output aliases.json

# Custom tickers
python ticker_aliases.py --tickers EXC XEL AEP CEG MSFT GOOG NVDA --output aliases.json

# From a file (one ticker per line)
python ticker_aliases.py --from-file my_tickers.txt --output aliases.json

# Add extra manual aliases
python ticker_aliases.py --tickers MSFT AAPL \
  --extra-aliases "MSFT:Azure|Windows;AAPL:iPhone|Mac" \
  --output aliases.json
```

You can then load it in `news_harm.py`. If `aliases.json` is missing, the script will attempt to build one at runtime.

---

## What the outputs contain

### `news_db.xlsx` (persistent database)
- **Columns**: Date, Source, Title, Link, Sentiment, ID, Tickers, Companies
- Behavior: appends new rows, **de‑dupes by ID**, keeps the newest date on conflicts.
- Formatting: Date (80 px; `YYYY‑MM‑DD`), Source (220 px), thin row borders.

### `news_last7.xlsx` (rolling last N days)
- **Columns**: Date, Source, Title, Summary, Link, Sentiment, ID, Tickers, Companies
- Behavior: shows items **within the last `--days`** from **today**.
- Formatting: Summary (500 px), Source (220 px), Date (80 px), thin row borders.

### `news_outputs.xlsx` (legacy multi‑sheet)
- **Sheets**:
  - **RawNews**: unique articles (last N days)
  - **MappedScored**: one row per (article × ticker) with sentiment
  - **DailySignals**: per‑ticker daily mean sentiment, article count, and a simple BUY/HOLD/SELL tag
- Date shows as `YYYY‑MM‑DD` and the Date column is 80 px wide on all sheets.

---

## How it works (pipeline)

1. **Fetch feeds** (global + Argentina + Yahoo per‑ticker RSS). fileciteturn3file1  
2. **Normalize & de‑dupe** by a stable **ID** (MD5 of title+summary+link).  
3. **Annotate companies**: search headline+summary for aliases and add **Tickers/Companies** columns.  
4. **Score sentiment** once per unique article (VADER/FinBERT).  
5. **Write DB** (`news_db.xlsx`) – append + de‑dupe by ID.  
6. **Write Last‑N view** (`news_last7.xlsx`) – filter by `--days`, include `Summary`.  
7. **Legacy sheets** – explode per‑ticker, aggregate daily, (optional) join price/returns. fileciteturn3file1

---

## Customization

- **Tickers**: pass `--tickers` or rebuild `aliases.json` with `ticker_aliases.py`.
- **Feeds**: edit the `AR_FEEDS`, `GLOBAL_FEEDS`, and ticker RSS template near the top of the script. fileciteturn3file1
- **Formatting**: change the pixel widths (Date 80, Source 220, Summary 500) in the Excel writer helpers.  
- **Sentiment thresholds** (BUY/HOLD/SELL) & lookahead returns are configurable in `aggregate_daily` / `add_returns`. fileciteturn3file1

---

## Scheduling (optional)

- **Windows Task Scheduler**: run the command daily to grow your DB.
- **cron (Linux/macOS)**: `0 8 * * 1-5 python /path/news_harm.py --tickers EXC XEL MSFT --days 7`

---

## Troubleshooting

- **No module named news_harm**: Run from the folder that contains the file: `py .\news_harm.py ...`  
- **VADER lexicon not found**: The script downloads it automatically; ensure your environment can reach `nltk` mirrors.  
- **FinBERT is slow or fails**: Use `--backend vader` or pre‑download the model.  
- **Dates show `00:00:00`**: The script now writes pure dates and applies `YYYY‑MM‑DD` formatting.  
- **Columns look squeezed**: Widths are set automatically; adjust pixel values if you prefer.  
- **Python 3.13 issues**: Try Python 3.12.

---

## Credits

- RSS parsing via **feedparser**
- Sentiment via **NLTK VADER** (default) and optional **FinBERT**
- Market data via **yfinance** (optional)
- Excel writing via **openpyxl**

---

## License

This project is provided as‑is, without warranty. You may adapt it freely within your own projects.
