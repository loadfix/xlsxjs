#!/usr/bin/env python
"""Generate tests/render-test/comments/workbook.xlsx.

Exercises:
  - two classic (non-threaded) comments, each with a distinct author
  - anchors at A1 and B2 so the renderer can prove it indexes by ref

Run:
    ~/code/python-xlsx/.venv/bin/python scripts/make-comments-fixture.py
"""

from __future__ import annotations

from pathlib import Path

from xlsx import Workbook
from xlsx.comments import Comment

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
OUT_DIR = REPO / "tests" / "render-test" / "comments"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "workbook.xlsx"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Notes"

    ws["A1"] = "hello"
    ws["A1"].comment = Comment("look here", "Alice")

    ws["B2"] = 42
    ws["B2"].comment = Comment("total count", "Bob")

    wb.save(str(OUT_PATH))
    print(f"wrote {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
