# pdf2pptx

**Utility to convert a PDF slideshow to PowerPoint PPTX.**

Renders each page of a PDF as a PNG image and builds a PowerPoint slideshow
from those images. Useful when you want to use PowerPoint to present a set of
PDF slides (e.g. slides from Beamer). You can then use the presentation
capabilities of PowerPoint — notes, ink on slides, laser pointer, presenter
view — with slides authored in LaTeX.

The original author uses this to present PDF slides on a Surface Pro, annotating
them with a stylus while presenting.

Uses [PyMuPDF](https://github.com/pymupdf/PyMuPDF) and
[python-pptx](https://github.com/scanny/python-pptx) to do the hard work.

> **This is a fork** of [kevinmcguinness/pdf2pptx](https://github.com/kevinmcguinness/pdf2pptx),
> updated to work with current versions of its dependencies and with modern
> Python packaging. See [Differences from upstream](#differences-from-upstream).

## What it does

- Renders every PDF page at a configurable DPI (default 300).
- Sets the slide aspect ratio to match the first page of the PDF, so 16:9
  Beamer slides produce a 16:9 deck with no letterboxing.
- Places each rendered page as a full-bleed picture on a blank slide layout.
- Optionally converts only a page range.

Note that the result is **images, not editable text** — this is a presentation
wrapper around a PDF, not a PDF-to-PowerPoint content converter. Text in the
output is not selectable or searchable.

## Requirements

- Python 3.8 or newer
- PyMuPDF >= 1.20.1
- python-pptx >= 0.6.21
- click >= 8.0
- tqdm >= 4.0

Dependencies are installed automatically by every method below.

## Installation

Pick whichever fits how you work. If you just want the command available and
don't care about the details, use **Option 1**.

### Option 1 — Install this fork directly from GitHub

```bash
pip install git+https://github.com/niravlekinwala/pdf2pptx.git
```

This is the recommended route. It installs the current code with relaxed
dependency bounds, so it works with an up-to-date PyMuPDF.

### Option 2 — Clone and install

Useful if you want the source around to read or modify.

```bash
git clone https://github.com/niravlekinwala/pdf2pptx.git
cd pdf2pptx
pip install .
```

### Option 3 — Editable / development install

Changes to the source take effect immediately, without reinstalling.

```bash
git clone https://github.com/niravlekinwala/pdf2pptx.git
cd pdf2pptx
pip install -e .
```

### Option 4 — Isolated CLI install with pipx

Best if you want `pdf2pptx` on your PATH without adding its dependencies to any
project environment. pipx gives the tool its own private virtualenv.

```bash
# install pipx first if you don't have it
python -m pip install --user pipx
python -m pipx ensurepath

pipx install git+https://github.com/niravlekinwala/pdf2pptx.git
```

### Option 5 — uv

If you use [uv](https://github.com/astral-sh/uv):

```bash
# install as a persistent tool
uv tool install git+https://github.com/niravlekinwala/pdf2pptx.git

# or run it once without installing at all
uv tool run --from git+https://github.com/niravlekinwala/pdf2pptx.git pdf2pptx slides.pdf
```

### Option 6 — Dedicated conda environment

Keeps the tool and its dependencies out of your `base` environment, which is
generally a good idea.

```bash
conda create -n pdf2pptx python=3.12
conda activate pdf2pptx
pip install git+https://github.com/niravlekinwala/pdf2pptx.git
```

To use it later, `conda activate pdf2pptx` first. To remove it entirely:
`conda env remove -n pdf2pptx`.

### Option 7 — Install from PyPI (upstream, not this fork)

```bash
pip install pdf2pptx
```

⚠️ The package on PyPI is upstream version 1.0.5, which **hard-pins
`pymupdf==1.20.1`**. On a modern environment this will either downgrade your
PyMuPDF or fail to build, since 1.20.1 has no wheels for recent Python
versions. See [Troubleshooting](#troubleshooting). This fork is not published
to PyPI — use one of the options above instead.

### Verifying the install

```bash
pdf2pptx --help
```

If the command isn't found, the install directory may not be on your `PATH`.
You can always invoke it as a module instead:

```bash
python -m pdf2pptx.cli --help
```

## Usage

Convert a whole PDF. This creates `slides.pptx` next to `slides.pdf`:

```bash
pdf2pptx slides.pdf
```

### Options

| Option | Default | Description |
| --- | --- | --- |
| `-o`, `--output` | `PDF_FILE.pptx` | Where to write the pptx |
| `-r`, `--resolution` | `300` | Render resolution in DPI |
| `-q`, `--quiet` | off | Suppress the progress bar and info output |
| `--from` | `0` | First PDF page to copy (**0-based**) |
| `--count` | all pages | Number of pages to copy |
| `--help` | | Show help and exit |

### Examples

Write to a specific file:

```bash
pdf2pptx -o ~/talks/final.pptx slides.pdf
```

Lower the resolution to keep the file size down — useful for long decks, since
300 DPI on a 16:9 slide is a 3000×1688 px image per page:

```bash
pdf2pptx -r 150 slides.pdf
```

Convert only pages 6–10. `--from` is 0-based, so page 6 as displayed in a PDF
reader is index 5:

```bash
pdf2pptx --from 5 --count 5 slides.pdf
```

Silence output, e.g. in a script:

```bash
pdf2pptx -q slides.pdf
```

Convert every PDF in the current directory:

```bash
for f in *.pdf; do pdf2pptx "$f"; done
```

### Use as a library

```python
from pdf2pptx import convert_pdf2pptx

convert_pdf2pptx(
    'slides.pdf',
    'slides.pptx',
    resolution=300,
    start_page=0,
    page_count=None,
    quiet=False,
)
```

Passing `output_file=None` writes alongside the input with a `.pptx` suffix.
Passing `page_count=None` converts through to the end of the document.

## Choosing a resolution

Higher DPI means sharper slides and larger files. Rendered pixel width is
`page_width_in_points / 72 * dpi`.

| DPI | 16:9 slide (720×405 pt) | Rough use |
| --- | --- | --- |
| 96 | 960×540 | Drafts, quick previews |
| 150 | 1500×844 | Screen presentation, small files |
| 300 | 3000×1688 | Default; good for projectors and zooming |
| 600 | 6000×3375 | Print or very large displays; big files |

## Troubleshooting

**`pip install pdf2pptx` downgrades PyMuPDF, or fails to build it.**
That's the PyPI release (1.0.5) enforcing `pymupdf==1.20.1`. That version has
no wheels for recent Python releases, so pip tries to build it from source and
fails. Install this fork instead, which uses `pymupdf>=1.20.1`.

**`ModuleNotFoundError: No module named 'fitz'`**
Older PyMuPDF exposed itself as `fitz`; from 1.24.3 the top-level name is
`pymupdf`. This fork imports `pymupdf` and falls back to `fitz`, so it works
either way. If you hit this, PyMuPDF probably isn't installed at all:
`pip install pymupdf`.

**`pdf2pptx: command not found` after installing.**
The scripts directory isn't on your `PATH`. Use `python -m pdf2pptx.cli`, or
install with pipx (Option 4), which handles `PATH` for you.

**The output deck is enormous.**
Lower the DPI with `-r`. 150 DPI is usually indistinguishable when projected
and roughly a quarter the pixels of the default.

**Slides are letterboxed or cropped in PowerPoint.**
The deck's aspect ratio is taken from the *first* page of the PDF. If your PDF
mixes page sizes, split it or convert the differently-sized pages separately.

## How it works

For each requested page, PyMuPDF renders a pixmap at `dpi/72` zoom, which is
encoded as PNG in memory. python-pptx then creates a slide using the blank
layout (index 6 of the default template) and adds the PNG as a picture scaled
to the full slide height at position (0, 0). The deck's `slide_width` is set
from the first page's aspect ratio before any slides are added. Nothing touches
the disk except the final `.pptx`.

## Differences from upstream

This fork is functionally identical in output; the changes are about running on
a current Python stack.

- **Relaxed the `pymupdf==1.20.1` pin to `pymupdf>=1.20.1`.** The hard pin was
  the reason the package couldn't be installed on recent Python versions.
- **Added lower bounds** for `python-pptx`, `click`, and `tqdm` instead of
  leaving them unconstrained.
- **Switched `import fitz` to `import pymupdf`**, with a fallback to `fitz` for
  PyMuPDF older than 1.24.3. `fitz` is the legacy alias and is slated for
  removal upstream.
- **Migrated packaging from `setup.py` to `pyproject.toml`** (PEP 517/621).
  `pip install -e .` under `setup.py develop` is deprecated and stops working
  in pip 25.3; the new layout builds a PEP 660 editable wheel instead.
- **Declared `requires-python = ">=3.8"`** and refreshed the trove classifiers,
  which previously listed only Python 3.6.
- Version bumped to 1.0.6.

Verified against Python 3.13.5, PyMuPDF 1.26.3, python-pptx 1.0.2, click 8.1.8,
and tqdm 4.67.1.

## Credits

Original work by [Kevin McGuinness](https://github.com/kevinmcguinness) —
[kevinmcguinness/pdf2pptx](https://github.com/kevinmcguinness/pdf2pptx).
This fork only modernizes packaging and dependency handling; the conversion
logic is his.

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2020 Kevin McGuinness.
