"""Usage: python -m src.generate_ppt <slides_root> <output.pptx>

Generic PPT generator driver. Slide-specific rendering can be implemented in
per-slide modules (for example `src.slide1`) which export a `process(prs, folder)`
function. This file provides reusable helpers and a small, clear control flow.
"""

from pathlib import Path
import sys
from typing import Optional

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image

# Import optional slide-specific handlers
try:
    from . import slide1
except Exception:
    slide1 = None
try:
    from . import slide2
except Exception:
    slide2 = None


def _read_meta_title(folder: Path) -> Optional[str]:
    meta = folder / 'meta.txt'
    if not meta.exists():
        return None
    try:
        return meta.read_text(encoding='utf-8').strip().splitlines()[0]
    except Exception:
        return None


def add_title_slide(prs: Presentation, title_text: str, subtitle_text: Optional[str] = None):
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    # some templates may not have a subtitle placeholder
    try:
        subtitle = slide.placeholders[1]
    except Exception:
        subtitle = None
    if subtitle_text and subtitle is not None:
        subtitle.text = subtitle_text
    return slide


def add_table_slide(prs: Presentation, title: str, df: pd.DataFrame, *, left=Inches(0.5), top=Inches(0.5), width=Inches(9)):
    blank = prs.slide_layouts[5]
    slide = prs.slides.add_slide(blank)
    # title textbox
    tx = slide.shapes.add_textbox(left, Inches(0.2), width, Inches(0.5)).text_frame
    p = tx.paragraphs[0]
    p.text = title
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = PP_ALIGN.LEFT

    rows, cols = df.shape[0] + 1, df.shape[1]
    table = slide.shapes.add_table(rows, cols, left, top + Inches(0.6), width, Inches(4)).table

    # header
    for j, col in enumerate(df.columns):
        cell = table.cell(0, j)
        cell.text = str(col)
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True

    # body
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, val in enumerate(row):
            table.cell(i, j).text = str(val)
    return slide


def _px_to_inches(px: int, dpi: int = 96) -> float:
    return px / dpi


def add_images_from_folder(slide, folder: Path, *, max_width=Inches(6), left=Inches(1), top=Inches(1)):
    imgs = sorted(folder.glob('*.png')) + sorted(folder.glob('*.jpg')) + sorted(folder.glob('*.jpeg'))
    y = top
    for img in imgs:
        try:
            im = Image.open(img)
            w_px, h_px = im.size
            w_in = _px_to_inches(w_px)
            h_in = _px_to_inches(h_px)
            scale = min(max_width.inches / w_in, 1.0) if w_in > 0 else 1.0
            slide.shapes.add_picture(str(img), left, y, width=Inches(w_in * scale))
            y = y + Inches(h_in * scale) + Inches(0.2)
        except Exception:
            # skip unreadable images
            continue


def process_generic_folder(prs: Presentation, folder: Path):
    """Fallback processor: add a title, render any CSVs as tables, then images."""
    title = _read_meta_title(folder) or folder.name
    add_title_slide(prs, title)

    for csv in sorted(folder.glob('*.csv')):
        try:
            df = pd.read_csv(csv)
            add_table_slide(prs, csv.stem, df)
        except Exception:
            continue

    imgs = list(folder.glob('*.png')) + list(folder.glob('*.jpg')) + list(folder.glob('*.jpeg'))
    if imgs:
        blank = prs.slide_layouts[5]
        slide = prs.slides.add_slide(blank)
        add_images_from_folder(slide, folder)


def generate(slides_root: str, output: str):
    prs = Presentation()
    slides_root = Path(slides_root)
    if not slides_root.exists():
        raise SystemExit(f"Slides root {slides_root} does not exist")

    for folder in sorted([p for p in slides_root.iterdir() if p.is_dir()]):
        # slide-specific handlers take precedence
        if folder.name == 'slide2' and slide2 is not None:
            try:
                slide2.process(prs, folder)
                continue
            except Exception:
                pass

        if folder.name == 'slide1' and slide1 is not None:
            try:
                slide1.process(prs, folder)
                continue
            except Exception:
                pass

        process_generic_folder(prs, folder)

    prs.save(output)
    print(f"Saved presentation to {output}")


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        raise SystemExit("Usage: python -m src.generate_ppt <slides_root> <output.pptx>")
    generate(argv[1], argv[2])


if __name__ == '__main__':
    main(sys.argv)
