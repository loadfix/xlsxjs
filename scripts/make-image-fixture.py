#!/usr/bin/env python
"""Generate tests/render-test/image/workbook.xlsx with a small embedded PNG
(generated via PIL, no external asset) anchored at cell B2.

Run:
    ~/code/python-xlsx/.venv/bin/python scripts/make-image-fixture.py
"""

from __future__ import annotations

import io
from pathlib import Path

from PIL import Image as PImg

from xlsx import Workbook
from xlsx.drawing.image import Image

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
OUT_DIR = REPO / "tests" / "render-test" / "image"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "workbook.xlsx"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "WithImage"

    ws["A1"] = "title"
    ws["A2"] = "logo below:"

    # Procedural 32×32 red square so the fixture has no external asset.
    # python-xlsx's Image.open reads lazily, so write to disk first.
    pil = PImg.new("RGB", (32, 32), (255, 0, 0))
    png_path = OUT_DIR / "red.png"
    pil.save(png_path, format="PNG")
    img = Image(str(png_path))
    img.anchor = "B3"
    ws.add_image(img)

    wb.save(str(OUT_PATH))
    print(f"wrote {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
