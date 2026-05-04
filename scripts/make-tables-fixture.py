#!/usr/bin/env python
"""Generate tests/render-test/tables/workbook.xlsx.

Exercises:
  - frozen panes (top row + first column)
  - autoFilter range
  - a defined named table

Run:
    ~/code/python-xlsx/.venv/bin/python scripts/make-tables-fixture.py
"""

from __future__ import annotations

from pathlib import Path

from xlsx import Workbook
from xlsx.worksheet.table import Table

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
OUT_DIR = REPO / "tests" / "render-test" / "tables"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "workbook.xlsx"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory"

    # Header + data.
    ws["A1"] = "SKU"; ws["B1"] = "Name"; ws["C1"] = "Region"; ws["D1"] = "Stock"
    rows = [
        ("A-100", "Widget",  "North",  42),
        ("A-101", "Gadget",  "South", 118),
        ("A-102", "Sprocket","East",    7),
        ("A-103", "Gizmo",   "West",   55),
    ]
    for i, (sku, nm, reg, stock) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=sku)
        ws.cell(row=i, column=2, value=nm)
        ws.cell(row=i, column=3, value=reg)
        ws.cell(row=i, column=4, value=stock)

    # Freeze header row + first column.
    ws.freeze_panes = "B2"

    # Auto-filter spanning the full data range.
    ws.auto_filter.ref = "A1:D5"

    # Defined table (named).
    table = Table(displayName="InventoryTable", name="InventoryTable",
                  ref="A1:D5")
    ws.add_table(table)

    wb.save(str(OUT_PATH))
    print(f"wrote {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
