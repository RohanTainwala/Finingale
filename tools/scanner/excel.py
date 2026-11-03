#!/usr/bin/env python3
"""Build the downloadable Excel file from a scan's JSON (US or India).

    python3 excel.py data/daily-scan.json data/daily-scan.xlsx
    python3 excel.py data/daily-scan-india.json data/daily-scan-india.xlsx

Needs openpyxl (the daily workflow installs it).
"""
import datetime
import json
import sys
import urllib.parse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SETUPS = [  # key, column title, fill
    ("breakout_20d", "20-day breakout", "FDF1CC"),
    ("bull_flag", "Bull flag", "E3F1FA"),
    ("reclaim_ma20", "Reclaimed 20-day avg", "E3F4E8"),
    ("oversold", "Oversold", "FBE4DE"),
    ("oversold_bounce", "Oversold bounce", "EFE6FB"),
]
RULES = {
    "breakout_20d": "The latest close is above every close of the 20 sessions before it.",
    "bull_flag": "Up more than 10% over 10 sessions, then the last 5 closes held inside a 5% range.",
    "reclaim_ma20": "Yesterday's close was below its 20-day moving average; the latest close is above it.",
    "oversold": "The 14-day Relative Strength Index (Wilder) is below 30.",
    "oversold_bounce": "RSI was below 30 the session before, and the latest close was higher.",
}
DISCLAIMER = ("General information only. Produced automatically by a rules-based screen. It does not take into "
              "account your objectives, financial situation or needs and is not a recommendation to buy or sell "
              "any security. Prices are previous closes from free, unofficial data (Yahoo Finance) and may be "
              "delayed or wrong. Past patterns do not predict future returns.")

EVERGREEN, FERN, INK2 = "0F3D2E", "1E7A4C", "4A6358"
HEAD_FILL = PatternFill("solid", fgColor=EVERGREEN)
HEAD_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, color=EVERGREEN)
MUTED_FONT = Font(size=10, color=INK2)
WRAP = Alignment(wrap_text=True, vertical="top")


def close_label(as_of):
    d = datetime.date.fromisoformat(as_of)
    return d.strftime("%A %-d %B %Y")


def is_india(doc):
    return doc.get("market") == "india"


def chart_url(sym, india):
    # Yahoo spells US class shares with a dash (BRK-B); Indian symbols keep .NS / .BO
    return "https://finance.yahoo.com/quote/" + urllib.parse.quote(sym if india else sym.replace(".", "-"))


def columns(doc):
    """(title, width, value(h), number format) for each column after the setup flags."""
    india = is_india(doc)
    money = '"\u20b9"#,##0.00' if india else "$#,##0.00"
    big = '"\u20b9"#,##0' if india else "$#,##0"
    traded = lambda h: h["p"] * h["v"] if h.get("v") is not None else None
    cols = [
        ("Last close (INR)" if india else "Last close (USD)", 12, lambda h: h["p"], money),
        ("RSI 14", 8, lambda h: h.get("r"), "0.0"),
        ("Volume (shares)", 15, lambda h: h.get("v"), "#,##0"),
        ("Rupees traded (INR)" if india else "Dollars traded (USD)", 18, traded, big),
    ]
    if india:
        cols.append(("Est. market cap (Rs crore)", 16, lambda h: h.get("m"), "#,##0"))
    return cols


def matches_sheet(ws, doc):
    hits = doc["hits"]
    india = is_india(doc)
    ws.title = "Matches"
    ws["A1"] = "Finingale daily stock scanner" + (": India (NSE and BSE)" if india else ": US")
    ws["A1"].font = TITLE_FONT
    if india:
        ws["A2"] = "NSE/BSE close of %s · %s matches from %s stocks passing the filters (%s listed companies checked) · finingale.com/daily-scanner.html#india" % (
            close_label(doc["as_of"]), f"{len(hits):,}", f"{doc['scanned']:,}", f"{doc['universe']:,}")
    else:
        ws["A2"] = "US close of %s · %s matches from %s liquid stocks (%s checked) · finingale.com/daily-scanner.html" % (
            close_label(doc["as_of"]), f"{len(hits):,}", f"{doc['scanned']:,}", f"{doc['universe']:,}")
    ws["A2"].font = MUTED_FONT

    lead = [("Ticker", 13 if india else 9, lambda h: h["s"].rsplit(".", 1)[0] if india else h["s"], None)]
    if india:
        lead.append(("Exchange", 9, lambda h: "BSE" if h["s"].endswith(".BO") else "NSE", None))
    lead += [("Company", 32, lambda h: h["n"], None),
             ("Setups", 30, lambda h: ", ".join(labels.get(k, k) for k in h["k"]), None)]
    labels = {k: t for k, t, _ in SETUPS}
    flags = [(t, 11, None, None) for _, t, _ in SETUPS]
    tail = columns(doc) + [("Chart", 12, None, None)]
    spec = lead + flags + tail
    first_flag = len(lead) + 1

    ws["A3"] = DISCLAIMER + (" Market caps are estimates from AMFI's six-monthly averages." if india else "")
    ws["A3"].font = MUTED_FONT
    ws["A3"].alignment = WRAP
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(spec))
    ws.row_dimensions[3].height = 42

    top = 5
    for c, (title, width, _, _) in enumerate(spec, 1):
        cell = ws.cell(top, c, title)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[top].height = 30

    for i, h in enumerate(hits, top + 1):
        for c, (_, _, fn, fmt) in enumerate(spec, 1):
            if fn is None:  # setup flags and the chart link are filled below
                continue
            cell = ws.cell(i, c, fn(h))
            if fmt:
                cell.number_format = fmt
        for c, (k, _, fill) in enumerate(SETUPS, first_flag):
            if k in h["k"]:
                ws.cell(i, c, "Yes").fill = PatternFill("solid", fgColor=fill)
                ws.cell(i, c).alignment = Alignment(horizontal="center")
        link = ws.cell(i, len(spec), "Open chart")
        link.hyperlink = chart_url(h["s"], india)
        link.font = Font(color=FERN, underline="single")

    last = top + max(1, len(hits))
    table = Table(displayName="Matches", ref="A%d:%s%d" % (top, get_column_letter(len(spec)), last))
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = ws.cell(top + 1, len(lead))  # keep ticker and company in view


def summary_sheet(ws, doc):
    ws.title = "Summary"
    ws["A1"] = "Summary"
    ws["A1"].font = TITLE_FONT
    if is_india(doc):
        rows = [
            ("NSE/BSE close", close_label(doc["as_of"])),
            ("Generated (UTC)", doc["generated_at"].replace("T", " ").replace("Z", "")),
            ("Listed companies checked", doc["universe"]),
            ("Passed the filters", doc["scanned"]),
            ("Price filter", "Rs %g or more" % doc["filters"]["min_price"]),
            ("Market-cap filter", "Rs %s crore (Rs 5 billion) or more, estimated" % f"{doc['filters']['min_mcap_cr']:,.0f}"),
            ("Market-cap source", doc.get("mcap_source", "")),
            ("Matched at least one setup", len(doc["hits"])),
        ]
    else:
        rows = [
            ("US close", close_label(doc["as_of"])),
            ("Generated (UTC)", doc["generated_at"].replace("T", " ").replace("Z", "")),
            ("US stocks checked", doc["universe"]),
            ("Liquid enough to scan", doc["scanned"]),
            ("Price filter", "$%g or more" % doc["filters"]["min_price"]),
            ("Volume filter", "%s+ shares in the last session" % f"{doc['filters']['min_volume']:,}"),
            ("Matched at least one setup", len(doc["hits"])),
        ]
    for r, (k, v) in enumerate(rows, 3):
        ws.cell(r, 1, k).font = Font(bold=True, color=EVERGREEN)
        ws.cell(r, 2, v)
        if isinstance(v, int):
            ws.cell(r, 2).number_format = "#,##0"
            ws.cell(r, 2).alignment = Alignment(horizontal="left")

    r = len(rows) + 5
    for c, title in enumerate(["Setup", "Matches", "Rule (daily closes only)"], 1):
        ws.cell(r, c, title).font = HEAD_FONT
        ws.cell(r, c).fill = HEAD_FILL
    for k, title, fill in SETUPS:
        r += 1
        ws.cell(r, 1, title).fill = PatternFill("solid", fgColor=fill)
        ws.cell(r, 2, doc["counts"].get(k, 0)).number_format = "#,##0"
        ws.cell(r, 3, RULES[k]).alignment = WRAP
    r += 2
    ws.cell(r, 1, "Backtest").font = Font(bold=True, color=EVERGREEN)
    ws.cell(r + 1, 1, "Tested Jan 2025 to Oct 2026 on about 1,500 liquid US stocks against the same-day average stock: "
                      "none of the setups beat it reliably. Oversold came out highest (about +0.7% over 10 days) but "
                      "beat the average only 49.7% of the time. Costs and delisted stocks were not included."
                      + (" The backtest covers US stocks only; these setups have not been tested on Indian stocks."
                         if is_india(doc) else ""))
    ws.cell(r + 3, 1, DISCLAIMER)
    for rr in (r + 1, r + 3):
        ws.cell(rr, 1).alignment = WRAP
        ws.cell(rr, 1).font = MUTED_FONT
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=3)
        ws.row_dimensions[rr].height = 62
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 70


def main(src, out):
    with open(src) as f:
        doc = json.load(f)
    wb = Workbook()
    matches_sheet(wb.active, doc)
    summary_sheet(wb.create_sheet(), doc)
    wb.save(out)
    print("Wrote %s (%d matches)" % (out, len(doc["hits"])))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
