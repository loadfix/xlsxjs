#!/usr/bin/env python
"""Generate tests/render-test/cf-graphical/workbook.xlsx.

Exercises the graphical conditional-formatting rules that the main
cf fixture (`conditional-format`) doesn't touch: a 3-colour scale across
a numeric range and a data bar across the same range.

Run:
    ~/code/python-xlsx/.venv/bin/python scripts/make-cf-graphical-fixture.py
"""

from __future__ import annotations

from pathlib import Path

from xlsx import Workbook
from xlsx.formatting.rule import ColorScaleRule, DataBarRule

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
OUT_DIR = REPO / "tests" / "render-test" / "cf-graphical"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "workbook.xlsx"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Graphical"

    # Column A: 5 rows of numeric values spanning 0..100 so colour scale /
    # data bar stops hit predictable thresholds.
    ws["A1"] = "Score"
    ws["B1"] = "Bar"
    values = [0, 25, 50, 75, 100]
    for i, v in enumerate(values, start=2):
        ws.cell(row=i, column=1, value=v)
        ws.cell(row=i, column=2, value=v)

    # Red → yellow → green colour scale on A2:A6.
    ws.conditional_formatting.add(
        "A2:A6",
        ColorScaleRule(
            start_type="min", start_color="FFFF0000",
            mid_type="percentile", mid_value=50, mid_color="FFFFFF00",
            end_type="max", end_color="FF00FF00",
        ),
    )

    # Data bar on B2:B6 in blue, default min/max.
    ws.conditional_formatting.add(
        "B2:B6",
        DataBarRule(
            start_type="min", end_type="max",
            color="FF4472C4", showValue=True,
        ),
    )

    wb.save(str(OUT_PATH))
    print(f"wrote {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
