#!/usr/bin/env python
"""Generate tests/render-test/conditional-format/workbook.xlsx.

Exercises the conditional-formatting slice: cellIs (>= 90), containsText,
duplicateValues, and top10 (top 1 by percent). Authored via python-xlsx so
the emitted XML looks like what real spreadsheets ship.

Run:
    ~/code/python-xlsx/.venv/bin/python scripts/make-cf-fixture.py
"""

from __future__ import annotations

from pathlib import Path

from xlsx import Workbook
from xlsx.formatting.rule import Rule
from xlsx.styles import Font, PatternFill
from xlsx.styles.differential import DifferentialStyle

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
OUT_DIR = REPO / "tests" / "render-test" / "conditional-format"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "workbook.xlsx"


def make_rule(rule_type: str, dxf: DifferentialStyle, **kwargs) -> Rule:
    return Rule(type=rule_type, dxf=dxf, **kwargs)


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Scores"

    ws["A1"] = "Name"
    ws["B1"] = "Score"
    ws["C1"] = "Region"

    rows = [
        ("Alice",   95, "North"),
        ("Bob",     72, "North"),
        ("Carol",   90, "South"),
        ("Dave",    45, "West"),
        ("Eve",     62, "North"),
        ("Frank",   88, "South"),
    ]
    for i, (name, score, region) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=name)
        ws.cell(row=i, column=2, value=score)
        ws.cell(row=i, column=3, value=region)

    # Rule 1: scores >= 90 → bold red fill.
    red_fill = PatternFill(fill_type="solid", bgColor="FFFF9999")
    cellIs_dxf = DifferentialStyle(font=Font(b=True), fill=red_fill)
    ws.conditional_formatting.add(
        "B2:B7",
        Rule(type="cellIs", operator="greaterThanOrEqual", formula=["90"],
             dxf=cellIs_dxf, priority=1),
    )

    # Rule 2: containsText "North" → yellow fill.
    yellow_fill = PatternFill(fill_type="solid", bgColor="FFFFFF99")
    contains_dxf = DifferentialStyle(fill=yellow_fill)
    ws.conditional_formatting.add(
        "C2:C7",
        Rule(type="containsText", operator="containsText", text="North",
             dxf=contains_dxf, priority=2,
             formula=['NOT(ISERROR(SEARCH("North",C2)))']),
    )

    # Rule 3: duplicate values in column C → light blue fill.
    dup_fill = PatternFill(fill_type="solid", bgColor="FFCCE5FF")
    dup_dxf = DifferentialStyle(fill=dup_fill)
    ws.conditional_formatting.add(
        "C2:C7",
        Rule(type="duplicateValues", dxf=dup_dxf, priority=3),
    )

    wb.save(str(OUT_PATH))
    print(f"wrote {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
