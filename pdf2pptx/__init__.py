# Copyright (c) 2020 Kevin McGuinness <kevin.mcguinness@gmail.com>
import io

try:
    import pymupdf
except ImportError:
    # PyMuPDF only exposed the top-level `pymupdf` name from 1.24.3 onwards
    import fitz as pymupdf

from tqdm import tqdm
from pptx import Presentation
from pathlib import Path


__all__ = [
    'convert_pdf2pptx', 'parse_page_spec', 'parse_slide_size',
    'DEFAULT_RESOLUTION', 'DEFAULT_QUALITY', 'IMAGE_FORMATS',
]

DEFAULT_RESOLUTION = 300
DEFAULT_QUALITY = 85
IMAGE_FORMATS = ('png', 'jpeg')

# blank layout in the default python-pptx template
_BLANK_LAYOUT = 6

_POINTS_PER_INCH = 72


def parse_page_spec(spec, total):
    """
    Parse a 1-based page specification into a list of 0-based page indices.

    Accepts comma separated singles and ranges, e.g. ``"1-5,8,12-"``. An open
    ended range runs to the last page; ``"-4"`` starts at the first page.
    Pages keep the order given and duplicates are dropped.
    """
    pages = []
    for chunk in spec.split(','):
        chunk = chunk.strip()
        if not chunk:
            continue
        try:
            if '-' in chunk:
                lo_text, _, hi_text = chunk.partition('-')
                lo = int(lo_text) if lo_text.strip() else 1
                hi = int(hi_text) if hi_text.strip() else total
            else:
                lo = hi = int(chunk)
        except ValueError:
            raise ValueError(
                f'invalid page range {chunk!r}: expected a number '
                f'or a range like 3-7') from None
        if lo < 1:
            raise ValueError(f'page numbers start at 1, got {lo}')
        if lo > hi:
            raise ValueError(f'page range {chunk!r} runs backwards')
        if hi > total:
            raise ValueError(
                f'page {hi} is out of range: the document has {total} pages')
        pages.extend(range(lo - 1, hi))

    if not pages:
        raise ValueError('no pages selected')

    seen = set()
    unique = []
    for page in pages:
        if page not in seen:
            seen.add(page)
            unique.append(page)
    return unique


def parse_slide_size(spec):
    """
    Parse a slide size spec into an aspect ratio.

    Returns ``None`` for ``"auto"``, meaning the ratio should be taken from the
    first converted page. Otherwise accepts ``W:H`` or ``WxH``, e.g. ``16:9``.
    """
    if spec is None:
        return None
    text = spec.strip().lower()
    if text in ('auto', ''):
        return None
    separator = ':' if ':' in text else 'x' if 'x' in text else None
    if separator is None:
        raise ValueError(
            f'invalid slide size {spec!r}: expected auto, or a ratio '
            f'like 16:9')
    width_text, _, height_text = text.partition(separator)
    try:
        width = float(width_text)
        height = float(height_text)
    except ValueError:
        raise ValueError(f'invalid slide size {spec!r}: not a number') from None
    if width <= 0 or height <= 0:
        raise ValueError(f'invalid slide size {spec!r}: must be positive')
    return width / height


def _open_document(pdf_file, password):
    doc = pymupdf.open(pdf_file)
    if doc.needs_pass:
        if password is None:
            raise ValueError(
                f'{pdf_file} is password protected: supply one with --password')
        if not doc.authenticate(password):
            raise ValueError(f'incorrect password for {pdf_file}')
    return doc


def _select_pages(doc, pages, start_page, page_count):
    total = doc.page_count
    if total == 0:
        raise ValueError('the document has no pages')

    if pages is not None:
        return parse_page_spec(pages, total), None

    start_page = 0 if start_page is None else start_page
    if start_page < 0:
        raise ValueError(f'--from must not be negative, got {start_page}')
    if start_page >= total:
        raise ValueError(
            f'--from {start_page} is past the last page: the document has '
            f'{total} pages (--from counts from 0)')

    available = total - start_page
    if page_count is None:
        wanted = available
    elif page_count <= 0:
        raise ValueError(f'--count must be positive, got {page_count}')
    else:
        wanted = page_count

    truncated = wanted - available if wanted > available else None
    return list(range(start_page, start_page + min(wanted, available))), truncated


def _resolve_format(image_format, quality):
    fmt = (image_format or 'png').strip().lower()
    if fmt == 'jpg':
        fmt = 'jpeg'
    if fmt not in IMAGE_FORMATS:
        raise ValueError(
            f'unsupported image format {image_format!r}: '
            f'choose from {", ".join(IMAGE_FORMATS)}')
    if not 1 <= quality <= 100:
        raise ValueError(f'--quality must be between 1 and 100, got {quality}')
    return fmt


def _add_fitted_picture(slide, prs, image_file, page_aspect):
    """Add the image centred on the slide, scaled to fit without cropping."""
    slide_aspect = prs.slide_width / prs.slide_height
    if page_aspect >= slide_aspect:
        width = prs.slide_width
        height = int(round(prs.slide_width / page_aspect))
    else:
        height = prs.slide_height
        width = int(round(prs.slide_height * page_aspect))
    left = int(round((prs.slide_width - width) / 2))
    top = int(round((prs.slide_height - height) / 2))
    slide.shapes.add_picture(image_file, left, top, width=width, height=height)


def convert_pdf2pptx(
        pdf_file, output_file=None, resolution=DEFAULT_RESOLUTION,
        start_page=0, page_count=None, quiet=False, *,
        pages=None, image_format='png', quality=DEFAULT_QUALITY,
        notes=False, password=None, slide_size='auto'):
    """
    Convert a PDF slideshow to a PowerPoint presentation.

    Each selected page is rendered to an image and placed on its own slide.
    Returns the path the presentation was written to.
    """
    if resolution <= 0:
        raise ValueError(f'--resolution must be positive, got {resolution}')

    fmt = _resolve_format(image_format, quality)
    aspect_override = parse_slide_size(slide_size)

    doc = _open_document(pdf_file, password)
    page_numbers, truncated = _select_pages(doc, pages, start_page, page_count)

    if not quiet:
        print(f'{pdf_file} contains {doc.page_count} slides, '
              f'converting {len(page_numbers)}')
        if truncated:
            print(f'  note: only {len(page_numbers)} pages were available, '
                  f'{truncated} fewer than requested')

    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[_BLANK_LAYOUT]

    # configure presentation aspect ratio
    first_page = doc.load_page(page_numbers[0])
    aspect_ratio = aspect_override or (
        first_page.rect.width / first_page.rect.height)
    prs.slide_width = int(prs.slide_height * aspect_ratio)

    # transformation matrix: slide to pixmap
    zoom = resolution / _POINTS_PER_INCH
    matrix = pymupdf.Matrix(zoom, zoom, 0)

    page_iter = page_numbers if quiet else tqdm(page_numbers, unit='page')

    for page_no in page_iter:
        page = doc.load_page(page_no)

        # write slide as a pixmap
        pixmap = page.get_pixmap(matrix=matrix)
        image_data = pixmap.tobytes(output=fmt, jpg_quality=quality)
        image_file = io.BytesIO(image_data)

        # add a slide
        slide = prs.slides.add_slide(blank_slide_layout)
        _add_fitted_picture(
            slide, prs, image_file, page.rect.width / page.rect.height)

        if notes:
            text = page.get_text().strip()
            if text:
                slide.notes_slide.notes_text_frame.text = text

    if output_file is None:
        output_file = Path(pdf_file).with_suffix('.pptx')

    # save presentation
    prs.save(output_file)
    return Path(output_file)
