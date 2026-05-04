#!/usr/bin/env python
"""Generate tests/render-test/cf-icons/workbook.xlsx — a 3-traffic-lights
iconSet across a column of values.

Run:
    ~/code/python-xlsx/.venv/bin/python scripts/make-icons-fixture.py
"""

from __future__ import annotations

from pathlib import Path

from xlsx import Workbook
from xlsx.formatting.rule import IconSetRule

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
OUT_DIR = REPO / "tests" / "render-test" / "cf-icons"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "workbook.xlsx"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Icons"

    ws["A1"] = "Score"
    for i, v in enumerate([10, 40, 60, 85, 100], start=2):
        ws.cell(row=i, column=1, value=v)

    # Traffic-light icon set; default thresholds (0/33/67 percent).
    rule = IconSetRule(
        icon_style="3TrafficLights1",
        type="percent",
        values=[0, 33, 67],
        showValue=True,
    )
    ws.conditional_formatting.add("A2:A6", rule)

    wb.save(str(OUT_PATH))
    print(f"wrote {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
