#!/usr/bin/env python3
"""
tickers_config.py
-----------------
Define DEFAULT_TICKERS and build TICKER_ALIASES automatically using Yahoo Finance (yfinance).

Usage examples:
  # 1) Save aliases for default tickers to JSON
  python tickers_config.py --output aliases.json

  # 2) Specify custom tickers
  python tickers_config.py --tickers EXC XEL AEP CEG MSFT GOOG NVDA --output aliases.json

  # 3) Read tickers from a text file (one per line)
  python tickers_config.py --from-file my_tickers.txt --output aliases.json

  # 4) Add extra manual aliases
  python tickers_config.py --tickers MSFT AAPL --extra-aliases "MSFT:Azure|Windows;AAPL:iPhone|Mac" --output aliases.json

You can then load the JSON in your main bot:
  with open("aliases.json", "r", encoding="utf-8") as f:
      TICKER_ALIASES = json.load(f)
"""

import argparse
import json
import re
from typing import Dict, List, Iterable, Optional

# yfinance for metadata
try:
    import yfinance as yf
except Exception as ex:
    raise SystemExit("Please install yfinance: pip install yfinance") from ex


# -----------------------------
# Defaults
# -----------------------------
DEFAULT_TICKERS: List[str] = ["EXC", "XEL", "AEP", "CEG", "MSFT", "GOOG", "AAPL", "AMZN", "NVDA"]

# -----------------------------
# Helpers to generate aliases
# -----------------------------
_CORP_SUFFIXES = [
    r",?\s+Inc\.?",
    r",?\s+Incorporated",
    r",?\s+Corporation",
    r",?\s+Corp\.?",
    r",?\s+Company",
    r",?\s+Co\.?",
    r",?\s+Ltd\.?",
    r",?\s+PLC",
    r",?\s+S\.?A\.?",
]

def _strip_corp_suffix(name: str) -> str:
    s = name.strip()
    for suf in _CORP_SUFFIXES:
        s = re.sub(suf + r"$", "", s, flags=re.IGNORECASE)
    return s.strip()

def _split_on_separators(name: str) -> List[str]:
    # Create variants by splitting on common separators
    parts = re.split(r"[/\-–—:|]", name)
    variants = [name.strip()]
    for p in parts:
        p = p.strip()
        if p and p not in variants:
            variants.append(p)
    return variants

def _dedupe_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out = []
    for it in items:
        key = it.lower()
        if key not in seen and it:
            seen.add(key)
            out.append(it)
    return out

def _safe_add(aliases: List[str], *candidates: Optional[str]) -> None:
    for c in candidates:
        if c and isinstance(c, str):
            c = c.strip()
            if c:
                aliases.append(c)

def _yfin_aliases(ticker: str) -> List[str]:
    """
    Pull reasonable name variants from yfinance .info metadata.
    """
    aliases: List[str] = []
    try:
        info = yf.Ticker(ticker).info
    except Exception:
        info = {}

    # Primary names
    long_name = info.get("longName") or info.get("shortName") or info.get("displayName")
    short_name = info.get("shortName")
    display = info.get("displayName")

    _safe_add(aliases, long_name, short_name, display)

    # Strip corporate suffixes (Inc, Corp, etc.)
    if long_name:
        base = _strip_corp_suffix(long_name)
        _safe_add(aliases, base)
        # split variants
        for v in _split_on_separators(base):
            _safe_add(aliases, v)

    # Always include ticker itself
    _safe_add(aliases, ticker)

    return _dedupe_keep_order(aliases)


def build_aliases(tickers: List[str], extra_aliases: Optional[Dict[str, List[str]]] = None) -> Dict[str, List[str]]:
    mapping: Dict[str, List[str]] = {}
    for t in tickers:
        base = _yfin_aliases(t)
        # merge extra aliases
        if extra_aliases and t in extra_aliases:
            base.extend(extra_aliases[t])
        mapping[t] = _dedupe_keep_order(base)
    return mapping


def parse_extra_aliases(expr: str) -> Dict[str, List[str]]:
    """
    Parse a compact string for extra aliases:
      "MSFT:Azure|Windows;AAPL:iPhone|Mac"
    """
    result: Dict[str, List[str]] = {}
    if not expr:
        return result
    for block in expr.split(";"):
        if not block.strip():
            continue
        if ":" not in block:
            continue
        tk, vals = block.split(":", 1)
        tk = tk.strip().upper()
        al = [v.strip() for v in vals.split("|") if v.strip()]
        if tk and al:
            result[tk] = al
    return result


# -----------------------------
# CLI
# -----------------------------
def main():
    ap = argparse.ArgumentParser(description="Build TICKER_ALIASES using Yahoo Finance")
    ap.add_argument("--tickers", nargs="+", default=None, help="List of tickers to process")
    ap.add_argument("--from-file", type=str, default=None, help="Path to a text file with one ticker per line")
    ap.add_argument("--extra-aliases", type=str, default=None,
                    help='Extra aliases string, e.g. "MSFT:Azure|Windows;AAPL:iPhone|Mac"')
    ap.add_argument("--output", type=str, default="aliases.json", help="Output JSON path")
    args, _ = ap.parse_known_args()  # notebook-friendly

    # Resolve tickers
    tickers: List[str] = []
    if args.tickers:
        tickers = [t.strip().upper() for t in args.tickers if t.strip()]
    elif args.from_file:
        with open(args.from_file, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip().upper()
                if s:
                    tickers.append(s)
    else:
        tickers = DEFAULT_TICKERS[:]  # copy

    extras = parse_extra_aliases(args.extra_aliases) if args.extra_aliases else None
    mapping = build_aliases(tickers, extras)

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    print(f"[ok] Wrote {len(mapping)} tickers to {args.output}")
    for t, al in mapping.items():
        preview = ", ".join(al[:5]) + (" ..." if len(al) > 5 else "")
        print(f"  {t}: {preview}")


if __name__ == "__main__":
    main()