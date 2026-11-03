#!/usr/bin/env python3
"""Daily scan of Indian equities (NSE and BSE) -> data/daily-scan-india.json.

    python3 india.py OUT.json [--cache tools/scanner/amfi-mcap.csv]

The universe and market caps come from AMFI's "Average Market Capitalization of
listed companies" list, which covers every company listed on NSE and BSE and is
published twice a year. Each company's market cap today is estimated by scaling
its AMFI six-month average by how far its price has moved since that window.
Prices are Yahoo Finance daily closes (.NS first, .BO when NSE has none). The
five setups are the same rules as the US scan (scan.py's find_setups).

Filters: last close at least Rs 10 and estimated market cap at least Rs 500 crore
(Rs 5 billion). Needs pdftotext (poppler-utils) to read a new AMFI PDF; without
it, or when AMFI can't be reached, the cached CSV is used.
"""
import csv
import datetime
import importlib.util
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("scan", os.path.join(HERE, "scan.py"))
scan = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scan)

MIN_PRICE = 10.0          # rupees
MIN_MCAP_CR = 500.0       # Rs 500 crore = Rs 5 billion
PREFILTER_CR = 300.0      # AMFI average needed to be checked at all (allows for a big rally since)
MIN_SCANNED = 800         # a normal day passes ~1,500; far fewer means Yahoo failed
WINDOW = 63               # closes fed to find_setups, same as the US daily scan
IST = datetime.timedelta(hours=5, minutes=30)
AMFI_URL = "https://portal.amfiindia.com/spages/AverageMarketCapitalization{day}{mon}{year}.pdf"
DEFAULT_CACHE = os.path.join(HERE, "amfi-mcap.csv")


# ---------------------------------------------------------------- AMFI list

def amfi_periods(today):
    """Half-year ends AMFI may have published, newest first (30 Jun / 31 Dec)."""
    out = []
    for year in (today.year, today.year - 1):
        for end in (datetime.date(year, 12, 31), datetime.date(year, 6, 30)):
            if end < today:
                out.append(end)
    return out[:3]


def parse_amfi(text):
    """pdftotext -layout output -> [{isin, name, bse, nse, mcap_cr}]. Rows can wrap, and a
    long name can run straight into its ISIN, so each record is rebuilt before parsing."""
    records, cur = [], None
    for line in text.split("\n"):
        if re.match(r"^\s*\d+\s+\S", line) and not re.match(r"^\s*\d+\s+(month|Avg)", line):
            if cur:
                records.append(cur)
            cur = line
        elif cur is not None and line.strip():
            cur += " " + line.strip()
    if cur:
        records.append(cur)

    rows = []
    for rec in records:
        m = re.match(r"^\s*\d+\s+(.*?)\s*(IN[A-Z0-9]{10})(.*)$", rec)
        if not m:
            continue
        name, isin, rest = m.group(1).strip(), m.group(2), m.group(3)
        cat = re.search(r"(Large|Mid|Small) Cap\s*$", rest)
        if not cat:
            continue
        toks = rest[:cat.start()].split()
        num = lambda t: re.fullmatch(r"[\d,]+\.\d+", t) is not None
        if not toks or not num(toks[-1]):
            continue
        avg = float(toks[-1].replace(",", ""))
        toks = toks[:-1]
        bse = nse = None
        i = 0
        if i < len(toks) and not num(toks[i]):          # BSE symbol (or "-")
            bse = None if toks[i] == "-" else toks[i]
            i += 1
            if bse and i < len(toks) and num(toks[i]):  # its market cap
                i += 1
        if i < len(toks) and not num(toks[i]):          # NSE symbol (or "-")
            nse = None if toks[i] == "-" else toks[i]
        rows.append({"isin": isin, "name": name, "bse": bse, "nse": nse, "mcap_cr": avg})
    return rows


def load_amfi(cache_path, today):
    """Newest AMFI list we can read, refreshing the cache when a newer one parses cleanly."""
    cached, cached_period = [], None
    if os.path.exists(cache_path):
        with open(cache_path) as f:
            reader = csv.DictReader(f)
            for r in reader:
                cached_period = r.pop("period")
                r["mcap_cr"] = float(r["mcap_cr"])
                r["bse"] = r["bse"] or None
                r["nse"] = r["nse"] or None
                cached.append(r)
    for end in amfi_periods(today):
        if cached_period and end.isoformat() <= cached_period:
            break
        url = AMFI_URL.format(day=end.day, mon=end.strftime("%b"), year=end.year)
        try:
            with tempfile.TemporaryDirectory() as tmp:
                pdf = os.path.join(tmp, "amfi.pdf")
                # curl, not urllib: AMFI's server drops Python's TLS handshake
                subprocess.run(["curl", "-sSfL", "-m", "180", "-A", scan.UA, "-o", pdf, url],
                               check=True, capture_output=True, timeout=200)
                text = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True,
                                      check=True, timeout=120).stdout.decode("utf-8", "replace")
        except Exception as e:  # not published yet, unreachable, or no pdftotext
            print("AMFI %s: %s" % (end, e))
            continue
        rows = parse_amfi(text)
        if len(rows) < 4000:  # the list has ~5,400 companies; far fewer means the layout changed
            print("AMFI %s: only %d rows parsed, keeping the cache" % (end, len(rows)))
            continue
        with open(cache_path, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["period", "isin", "name", "bse", "nse", "mcap_cr"])
            w.writeheader()
            for r in rows:
                w.writerow(dict(r, period=end.isoformat(), bse=r["bse"] or "", nse=r["nse"] or ""))
        print("AMFI list for the six months to %s: %d companies" % (end, len(rows)))
        return rows, end
    if not cached:
        sys.exit("No AMFI market-cap list available.")
    print("Using cached AMFI list for the six months to %s" % cached_period)
    return cached, datetime.date.fromisoformat(cached_period)


# ---------------------------------------------------------------- prices

def ist_date(ts):
    return (datetime.datetime.fromtimestamp(ts, datetime.timezone.utc) + IST).date()


def fetch(symbols):
    """Yahoo symbols -> {symbol: series}, 20 a request, 2 years of daily closes."""
    data = {}
    for k in range(0, len(symbols), scan.SPARK_MAX):
        batch = symbols[k:k + scan.SPARK_MAX]
        if k:
            scan.SLEEP(scan.PAUSE)
        path = "/v7/finance/spark?symbols=%s&range=2y&interval=1d" % ",".join(
            urllib.request.quote(s, safe="") for s in batch)
        try:
            data.update(scan.parse_spark(scan.yahoo_get(path), {s.upper(): s for s in batch}))
        except (scan.ScanError, ValueError) as e:
            print("  batch %d failed: %s" % (k // scan.SPARK_MAX + 1, e))
    return data


def closes(series, now_ist):
    """[(IST date, close)] of complete sessions: today's bar only once the market has closed."""
    out = {}
    # Yahoo often leaves the latest day's close blank for hours after NSE shuts; its meta then
    # already holds that session's last trade (stamped 15:29-15:30 IST), which is the close.
    meta = series.get("meta") or {}
    last_px, last_ts = meta.get("regularMarketPrice"), meta.get("regularMarketTime")
    last_trade = datetime.datetime.fromtimestamp(last_ts, datetime.timezone.utc).replace(tzinfo=None) + IST if last_ts else None
    for ts, c in zip(series["t"], series["c"]):
        d = ist_date(ts)
        if c is None and last_px is not None and last_trade and last_trade.date() == d and last_trade.time() >= datetime.time(15, 25):
            c = last_px
        if c is None:
            continue
        if d > now_ist.date() or (d == now_ist.date() and now_ist.time() < datetime.time(16, 0)):
            continue
        out[d] = float(c)
    return sorted(out.items())


# ---------------------------------------------------------------- scan

def main(out_path, cache_path):
    now_ist = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None) + IST
    companies, period_end = load_amfi(cache_path, now_ist.date())
    window_start = datetime.date(period_end.year, period_end.month - 5, 1)
    pool = [c for c in companies if c["mcap_cr"] >= PREFILTER_CR and (c["nse"] or c["bse"])]
    print("%d of %d companies have an AMFI average of Rs %d crore or more" % (len(pool), len(companies), PREFILTER_CR))

    # NSE first (AMFI's NSE column is sometimes blank for NSE-listed stocks, so try the BSE symbol there too)
    want = {c["isin"]: (c["nse"] or c["bse"]) + ".NS" for c in pool}
    data = fetch(sorted(set(want.values())))
    retry = {isin: c["bse"] + ".BO" for isin, c in ((c["isin"], c) for c in pool)
             if c["bse"] and len((data.get(want[isin]) or {}).get("t") or []) < WINDOW}
    if retry:
        data.update(fetch(sorted(set(retry.values()))))
        want.update({i: s for i, s in retry.items() if len((data.get(s) or {}).get("t") or []) >= WINDOW})

    hits, dates = [], {}
    counts = {"fetched": 0, "short": 0, "filtered": 0, "scanned": 0}
    for c in pool:
        sym = want[c["isin"]]
        s = data.get(sym)
        if not s:
            continue
        counts["fetched"] += 1
        bars = closes(s, now_ist)
        if len(bars) < WINDOW:
            counts["short"] += 1
            continue
        last = bars[-1][1]
        in_window = [p for d, p in bars if window_start <= d <= period_end]
        mcap = c["mcap_cr"] * (last / (sum(in_window) / len(in_window))) if in_window else c["mcap_cr"]
        if last < MIN_PRICE or mcap < MIN_MCAP_CR:
            counts["filtered"] += 1
            continue
        counts["scanned"] += 1
        dates[bars[-1][0]] = dates.get(bars[-1][0], 0) + 1
        tags, nums = scan.find_setups([p for _, p in bars[-WINDOW:]])
        if tags:
            vol = (s.get("meta") or {}).get("regularMarketVolume")
            name = c["name"].title() if c["name"].isupper() else c["name"]  # AMFI shouts some names
            hits.append({"s": sym, "n": name, "k": tags, "p": round(last, 2),
                         "r": round(nums["rsi"], 1) if nums["rsi"] is not None else None,
                         "v": int(vol) if vol is not None else None, "m": round(mcap)})

    if counts["scanned"] < MIN_SCANNED:
        sys.exit("Scan looks incomplete (scanned %d); not publishing." % counts["scanned"])
    hits.sort(key=lambda h: (-(h["p"] * (h["v"] or 0)), h["s"]))
    tally = {}
    for h in hits:
        for k in h["k"]:
            tally[k] = tally.get(k, 0) + 1
    as_of = max(dates, key=lambda d: (dates[d], d))
    doc = {
        "market": "india",
        "currency": "INR",
        "generated_at": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "as_of": as_of.isoformat(),
        "universe": len(companies),
        "scanned": counts["scanned"],
        "failed": len(pool) - counts["fetched"],
        "filters": {"min_price": MIN_PRICE, "min_mcap_cr": MIN_MCAP_CR},
        "mcap_source": "AMFI six-month average to %s, scaled by price since" % period_end.isoformat(),
        "counts": tally,
        "hits": hits,
    }
    with open(out_path, "w") as f:
        json.dump(doc, f, separators=(",", ":"))
    print("India scan, closes through %s: %d companies checked, %d passed the filters, %d matches -> %s"
          % (as_of, len(companies), counts["scanned"], len(hits), out_path))


if __name__ == "__main__":
    args = sys.argv[1:]
    cache = DEFAULT_CACHE
    if "--cache" in args:
        i = args.index("--cache")
        cache = args[i + 1]
        del args[i:i + 2]
    if len(args) != 1:
        sys.exit(__doc__)
    main(args[0], cache)
