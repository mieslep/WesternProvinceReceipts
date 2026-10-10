#!/usr/bin/env python3
"""Generate one PDF receipt per spreadsheet row for Scouting Ireland Western Province.

Usage:
    python generate_receipts.py Test_receipt.xlsx [-o receipts] [--show-event] [--sep _]

Spreadsheet columns used: Date, Ref:, Transaction Details, Pay-IN (and Code, with --show-event).
Rows with no Pay-IN amount are skipped.
Output file name: <ref><sep><transaction details><sep><pay-in amount>.pdf
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from reportlab.lib.pagesizes import A5
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

ORG = "Scouting Ireland Western Province"
GREEN = (0.0, 0.35, 0.25)


def clean(name: str) -> str:
    """Make text safe for use in a file name."""
    return re.sub(r'[\\/:*?"<>|\r\n]+', "-", str(name)).strip()


def read_rows(path: Path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip().rstrip(":").lower() if h else "" for h in next(rows)]
    col = {h: i for i, h in enumerate(header)}
    for need in ("date", "ref", "transaction details", "pay-in"):
        if need not in col:
            sys.exit(f"Missing column '{need}' in {path} (found: {header})")
    for r in rows:
        ref, amount = r[col["ref"]], r[col["pay-in"]]
        if ref in (None, "") or amount in (None, ""):
            continue
        yield {
            "date": r[col["date"]],
            "ref": str(ref).strip(),
            "name": str(r[col["transaction details"]] or "").strip(),
            "amount": float(amount),
            "event": str(r[col["code"]]).strip() if "code" in col and r[col["code"]] else "",
        }


def draw_receipt(path: Path, rec: dict, show_event: bool):
    w, h = A5
    c = canvas.Canvas(str(path), pagesize=A5)
    c.setTitle(f"Receipt {rec['ref']}")
    c.setAuthor(ORG)
    c.setFillColorRGB(*GREEN)
    c.rect(0, h - 28 * mm, w, 28 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(15 * mm, h - 15 * mm, ORG)
    c.setFont("Helvetica", 10)
    c.drawString(15 * mm, h - 22 * mm, "Official Receipt")

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 26)
    c.drawString(15 * mm, h - 48 * mm, "Receipt")
    c.setFont("Helvetica", 10)
    c.drawRightString(w - 15 * mm, h - 48 * mm, f"Ref: {rec['ref']}")

    d = rec["date"]
    date_s = d.strftime("%d/%m/%Y") if isinstance(d, datetime) else str(d or "")
    fields = [("Name", rec["name"]), ("Amount", f"€{rec['amount']:,.2f}")]
    if show_event and rec["event"]:
        fields.append(("Event", rec["event"]))
    fields.append(("Date", date_s))

    y = h - 68 * mm
    for label, value in fields:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(15 * mm, y, f"{label} –")
        c.setFont("Helvetica", 12)
        c.drawString(45 * mm, y, value)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.line(15 * mm, y - 3 * mm, w - 15 * mm, y - 3 * mm)
        y -= 14 * mm

    c.setFont("Helvetica-Oblique", 13)
    c.drawString(15 * mm, y - 6 * mm, "With Thanks")
    c.showPage()
    c.save()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spreadsheet", type=Path)
    ap.add_argument("-o", "--outdir", type=Path, default=Path("receipts"))
    ap.add_argument("--show-event", action="store_true", help="include the Code column as 'Event'")
    ap.add_argument("--sep", default="_", help="separator in file names (default '_'; ':' is not allowed on Windows)")
    a = ap.parse_args()

    a.outdir.mkdir(parents=True, exist_ok=True)
    n = 0
    for rec in read_rows(a.spreadsheet):
        fname = a.sep.join([clean(rec["ref"]), clean(rec["name"]), f"{rec['amount']:.2f}"]) + ".pdf"
        draw_receipt(a.outdir / fname, rec, a.show_event)
        print("Created", a.outdir / fname)
        n += 1
    print(f"{n} receipt(s) written to {a.outdir}/")


if __name__ == "__main__":
    main()
