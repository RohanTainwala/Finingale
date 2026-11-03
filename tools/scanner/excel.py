#!/usr/bin/env python3
"""Build the downloadable Excel file from data/daily-scan.json.

    python3 excel.py data/daily-scan.json data/daily-scan.xlsx

Needs openpyxl (the daily workflow installs it).
"""
import datetime
import json
import sys

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


def matches_sheet(ws, doc):
    hits = doc["hits"]
    ws.title = "Matches"
    ws["A1"] = "Finingale daily stock scanner"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "US close of %s · %s matches from %s liquid stocks (%s checked) · finingale.com/daily-scanner.html" % (
        close_label(doc["as_of"]), f"{len(hits):,}", f"{doc['scanned']:,}", f"{doc['universe']:,}")
    ws["A2"].font = MUTED_FONT
    ws["A3"] = DISCLAIMER
    ws["A3"].font = MUTED_FONT
    ws["A3"].alignment = WRAP
    ws.merge_cells("A3:M3")
    ws.row_dimensions[3].height = 42

    head = ["Ticker", "Company", "Setups"] + [t for _, t, _ in SETUPS] + [
        "Last close (USD)", "RSI 14", "Volume (shares)", "Dollars traded (USD)", "Chart"]
    top = 5
    for c, title in enumerate(head, 1):
        cell = ws.cell(top, c, title)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[top].height = 30
    labels = {k: t for k, t, _ in SETUPS}

    for i, h in enumerate(hits, top + 1):
        url = "https://finance.yahoo.com/quote/" + h["s"].replace(".", "-")
        row = [h["s"], h["n"], ", ".join(labels.get(k, k) for k in h["k"])]
        row += ["Yes" if k in h["k"] else "" for k, _, _ in SETUPS]
        traded = h["p"] * h["v"] if h.get("v") is not None else None
        row += [h["p"], h.get("r"), h.get("v"), traded, url]
        for c, v in enumerate(row, 1):
            ws.cell(i, c, v)
        for c, (k, _, fill) in enumerate(SETUPS, 4):
            if k in h["k"]:
                ws.cell(i, c).fill = PatternFill("solid", fgColor=fill)
                ws.cell(i, c).alignment = Alignment(horizontal="center")
        link = ws.cell(i, len(head))
        link.hyperlink = url
        link.value = "Open chart"
        link.font = Font(color=FERN, underline="single")
        ws.cell(i, 9).number_format = "$#,##0.00"
        ws.cell(i, 10).number_format = "0.0"
        ws.cell(i, 11).number_format = "#,##0"
        ws.cell(i, 12).number_format = "$#,##0"

    last = top + max(1, len(hits))
    table = Table(displayName="Matches", ref="A%d:%s%d" % (top, get_column_letter(len(head)), last))
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = ws.cell(top + 1, 3)
    for c, w in enumerate([9, 32, 30, 11, 10, 13, 10, 11, 12, 8, 15, 17, 12], 1):
        ws.column_dimensions[get_column_letter(c)].width = w


def summary_sheet(ws, doc):
    ws.title = "Summary"
    ws["A1"] = "Summary"
    ws["A1"].font = TITLE_FONT
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
                      "beat the average only 49.7% of the time. Costs and delisted stocks were not included.")
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
