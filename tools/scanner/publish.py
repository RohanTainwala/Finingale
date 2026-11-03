#!/usr/bin/env python3
"""Turn the scanner's trader/ output into data/daily-scan.json for daily-scanner.html.

    python3 publish.py WORKDIR/trader OUT.json

Fails (exit 1) instead of publishing when the scan looks broken, so the page
keeps showing the last good day rather than an empty or partial one.
"""
import datetime
import json
import os
import sys

MIN_SCANNED = 500   # a normal day scans ~1,500 liquid stocks; far fewer means Yahoo failed


def main(trader_dir, out_path):
    with open(os.path.join(trader_dir, "watchlist.json")) as f:
        wl = json.load(f)
    names = {}
    with open(os.path.join(trader_dir, "universe.txt")) as f:
        for line in f:
            sym, _, name = line.rstrip("\n").partition("|")
            if sym:
                names[sym] = name.split(" - ")[0].strip()  # "Intel Corporation - Common Stock" -> "Intel Corporation"

    if wl.get("mode") != "daily" or (wl.get("scanned") or 0) < MIN_SCANNED:
        sys.exit("Scan looks incomplete (scanned %s); not publishing." % wl.get("scanned"))

    counts = {}
    hits = []
    for h in wl.get("hits", []):
        for s in h["setups"]:
            counts[s] = counts.get(s, 0) + 1
        hits.append({"s": h["symbol"], "n": names.get(h["symbol"], ""), "k": h["setups"],
                     "p": h["price"], "r": h.get("rsi"), "v": h.get("volume")})

    doc = {
        "generated_at": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "as_of": (wl.get("as_of") or "").replace(" close", ""),
        "universe": wl.get("universe"),
        "scanned": wl.get("scanned"),
        "failed": wl.get("failed"),
        "filters": wl.get("filters"),
        "counts": counts,
        "hits": hits,
    }
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w") as f:
        json.dump(doc, f, separators=(",", ":"))
    print("Published %d matches from %d scanned stocks to %s" % (len(hits), doc["scanned"], out_path))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
