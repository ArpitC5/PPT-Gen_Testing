PPT-Gen Testing
=================

This project generates PowerPoint presentations from a folder-per-slide structure. Each slide folder contains CSVs, images, and other data used to populate a slide. Run the generator to produce `output.pptx`.

Setup
-----

Install dependencies into your Python environment:

```bash
pip install -r requirements.txt
```

Try it
------

Run the generator:

```bash
python -m src.generate_ppt slides output.pptx
```

Project layout
--------------

- `slides/slide1/` - example slide data (CSV, images)
- `src/generate_ppt.py` - main generator

Architecture diagram
--------------------

See `docs/project_flow.svg` for a visual overview of how the generator, slide folders, and slide-specific modules interact.

