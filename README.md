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
- Sets the slide aspect ratio to match the PDF page, so 16:9 Beamer slides
  produce a 16:9 deck with no letterboxing — or force a ratio with `--size`.
- Places each rendered page centred on a blank slide, scaled to fit without
  cropping, which keeps mixed-page-size PDFs intact.
- Converts any subset of pages with a 1-based range spec (`--pages 1-5,8,12-`).
- Optionally copies each page's text into the slide's speaker notes.
- Reads encrypted PDFs, converts several files in one run, and refuses to
  silently overwrite an existing deck.

Note that the slides are **images, not editable text** — this is a presentation
wrapper around a PDF, not a PDF-to-PowerPoint content converter. Slide text is
not selectable, though `--notes` gives you a text copy in the notes pane.

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
| `-o`, `--output` | `PDF_FILE.pptx` | Where to write the pptx (single input only) |
| `-d`, `--output-dir` | alongside input | Directory to write the pptx files into |
| `-r`, `--resolution`, `--dpi` | `300` | Render resolution in DPI |
| `-f`, `--format` | `png` | Slide image format: `png` or `jpeg` |
| `--quality` | `85` | JPEG quality, 1–100 (ignored for PNG) |
| `-p`, `--pages` | all | Pages to convert, **1-based**, e.g. `1-5,8,12-` |
| `-s`, `--size` | `auto` | Slide aspect: `auto`, or a ratio like `16:9`, `4:3` |
| `-n`, `--notes` | off | Copy page text into the slide's speaker notes |
| `--password` | | Password for an encrypted PDF |
| `-y`, `--overwrite` | off | Replace the output file if it already exists |
| `-q`, `--quiet` | off | Suppress the progress bar and info output |
| `--from` | `0` | *Deprecated.* First page, **0-based** — use `--pages` |
| `--count` | all | *Deprecated.* Number of pages — use `--pages` |
| `--version` | | Print the version and exit |
| `--help` | | Show help and exit |

### Examples

Write to a specific file:

```bash
pdf2pptx -o ~/talks/final.pptx slides.pdf
```

Convert a subset of pages. `--pages` is 1-based and matches what your PDF
reader shows, so this takes pages 1 to 5, page 8, and page 12 to the end:

```bash
pdf2pptx --pages 1-5,8,12- slides.pdf
```

Lower the resolution to keep the file size down — at 300 DPI a 16:9 slide is a
3000×1688 px image per page:

```bash
pdf2pptx -r 150 slides.pdf
```

Use JPEG for image-heavy decks (see [Choosing a format](#choosing-a-format)):

```bash
pdf2pptx -f jpeg --quality 80 slides.pdf
```

Copy the PDF's text into the speaker notes, so you can read your slide content
in presenter view:

```bash
pdf2pptx --notes slides.pdf
```

Force a 16:9 deck even if the PDF pages are not, letterboxing them to fit:

```bash
pdf2pptx --size 16:9 slides.pdf
```

Open an encrypted PDF:

```bash
pdf2pptx --password 'secret' slides.pdf
```

Convert a whole directory in one run, writing the results elsewhere:

```bash
pdf2pptx --output-dir ~/talks/pptx *.pdf
```

Re-run over existing output — without `-y` the tool refuses rather than
clobbering your previous deck:

```bash
pdf2pptx -y slides.pdf
```

### Use as a library

```python
from pdf2pptx import convert_pdf2pptx

convert_pdf2pptx(
    'slides.pdf',
    'slides.pptx',
    resolution=300,
    pages='1-5,8',      # 1-based spec; None converts everything
    image_format='jpeg',
    quality=80,
    notes=True,
    slide_size='16:9',  # or 'auto' to take it from the page
    password=None,
    quiet=True,
)
```

Returns the `Path` the deck was written to. `output_file=None` writes alongside
the input with a `.pptx` suffix. The older positional signature
`convert_pdf2pptx(pdf, out, resolution, start_page, page_count, quiet)` still
works unchanged, so existing scripts keep running.

Two helpers are exported if you want to validate input yourself:

```python
from pdf2pptx import parse_page_spec, parse_slide_size

parse_page_spec('1-3,8', 10)   # -> [0, 1, 2, 7]   (0-based indices)
parse_slide_size('16:9')       # -> 1.777...       (None for 'auto')
```

Both raise `ValueError` with a readable message on bad input.

## Choosing a resolution

Higher DPI means sharper slides and larger files. Rendered pixel width is
`page_width_in_points / 72 * dpi`.

| DPI | 16:9 slide (720×405 pt) | Rough use |
| --- | --- | --- |
| 96 | 960×540 | Drafts, quick previews |
| 150 | 1500×844 | Screen presentation, small files |
| 300 | 3000×1688 | Default; good for projectors and zooming |
| 600 | 6000×3375 | Print or very large displays; big files |

## Choosing a format

PNG is the default and is lossless. JPEG helps mainly on decks full of
photographs or gradients; on flat text and line art PNG compresses about as
well, so JPEG buys little and costs you artefacts around text.

Measured on this tool at 300 DPI, `--quality 85`:

| Deck content | PNG | JPEG | Saving |
| --- | --- | --- | --- |
| Full-page photographic image | 2931 KB | 1355 KB | 2.2× smaller |
| Plain text slides | 170 KB | 151 KB | 1.1× smaller |

Reach for `-f jpeg` when a PNG deck is too big to email; stay on PNG when the
slides are mostly text. Lowering `--dpi` is usually the more effective lever
either way.

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
and roughly a quarter the pixels of the default. If the slides are photographic,
`-f jpeg` helps too — see [Choosing a format](#choosing-a-format).

**Slides are letterboxed in PowerPoint.**
The deck's aspect ratio comes from the first converted page. If your PDF mixes
page sizes, later pages are centred and scaled to fit, which leaves bars around
them rather than cropping. Force a single ratio with `--size 16:9`, or convert
the differently-sized pages separately with `--pages`.

**`--pages 3` gave me the wrong page.**
`--pages` is 1-based, matching your PDF reader. The older `--from` is 0-based;
they are deliberately different, which is why `--from` is deprecated.

**`... already exists, use --overwrite to replace it`**
The tool won't silently replace an existing deck. Pass `-y` to allow it.

**`... is password protected: supply one with --password`**
The PDF is encrypted. Pass `--password 'yourpassword'`.

**`--notes` produced empty notes.**
The PDF has no text layer — typically because it's a scan, or because the
slides were exported as outlines/images. There is no text for the tool to copy.

## How it works

For each selected page, PyMuPDF renders a pixmap at `dpi/72` zoom, encoded in
memory as PNG or JPEG. python-pptx creates a slide from the blank layout
(index 6 of the default template) and adds that image centred, scaled to fit
the slide without cropping. The deck's `slide_width` is set from the first
selected page's aspect ratio — or from `--size` — before any slides are added,
so a deck whose pages all share one ratio comes out edge to edge with no bars.
With `--notes`, the page's text layer is copied into the slide's notes. Nothing
touches the disk except the final `.pptx`.

## Differences from upstream

### Compatibility fixes

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

### New features

- **`--pages`** with a 1-based spec (`1-5,8,12-`), replacing the 0-based
  `--from`/`--count` pair. Both old options still work.
- **`--format` / `--quality`** for JPEG output.
- **`--notes`** to copy page text into the slides' speaker notes.
- **`--size`** to force a slide aspect ratio instead of deriving it.
- **`--password`** for encrypted PDFs.
- **Multiple input files** in one run, with **`--output-dir`**.
- **`--overwrite`**, so an existing deck is never silently replaced.
- **`--dpi`** as an alias for `--resolution`, and **`--version`**.

### Fixes

- **Out-of-range page selections no longer raise a traceback.** Upstream,
  `--from 50` or `--count 999` on a short document raised an unhandled
  exception from `load_page`. Selections are now validated, `--count` clamps to
  the pages available with a note, and bad input produces a readable message.
- **Pages are centred and scaled to fit** rather than sized by height alone, so
  PDFs with mixed page sizes are no longer cropped horizontally.
- The slide aspect ratio comes from the first *selected* page rather than
  always page 1.

Verified against Python 3.13.5, PyMuPDF 1.26.3, python-pptx 1.0.2, click 8.1.8,
and tqdm 4.67.1.

## Credits

Original work by [Kevin McGuinness](https://github.com/kevinmcguinness) —
[kevinmcguinness/pdf2pptx](https://github.com/kevinmcguinness/pdf2pptx).
This fork only modernizes packaging and dependency handling; the conversion
logic is his.

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2020 Kevin McGuinness.
