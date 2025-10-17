"""Slide-2 specific rendering.

This is a placeholder module for Slide 2. Add `process(prs, folder: Path)` logic
here to render slide 2 when ready.
"""
from pathlib import Path


def process(prs, folder: Path):
    """Placeholder processor for slide2. Currently it creates a simple title slide
    reading `meta.txt` if present.
    """
    title = folder.name
    meta = folder / 'meta.txt'
    if meta.exists():
        try:
            title = meta.read_text(encoding='utf-8').strip().splitlines()[0]
        except Exception:
            pass

    # minimal: use the first slide layout to create a title slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    try:
        slide.shapes.title.text = title
    except Exception:
        pass
