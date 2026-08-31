#!/usr/bin/env python
# Copyright (c) 2020 Kevin McGuinness <kevin.mcguinness@gmail.com>

import click
import sys

from pathlib import Path

from . import (
    convert_pdf2pptx, DEFAULT_RESOLUTION, DEFAULT_QUALITY, IMAGE_FORMATS)


arg = click.argument
opt = click.option


def _output_path(pdf_path, output_file, output_dir):
    if output_file is not None:
        return Path(output_file)
    destination = Path(pdf_path).with_suffix('.pptx')
    if output_dir is not None:
        destination = Path(output_dir) / destination.name
    return destination


@click.command()
@click.version_option(package_name='pdf2pptx')
@opt('-o', '--output', 'output_file', default=None,
     type=click.Path(dir_okay=False),
     help='location to save the pptx (default: PDF_FILE.pptx)')
@opt('-d', '--output-dir', default=None,
     type=click.Path(file_okay=False),
     help='directory to save the pptx files in')
@opt('-r', '--resolution', '--dpi', 'resolution', default=DEFAULT_RESOLUTION,
     type=int, show_default=True,
     help='resolution in dots per inch')
@opt('-f', '--format', 'image_format', default='png',
     type=click.Choice(IMAGE_FORMATS, case_sensitive=False), show_default=True,
     help='image format used for the rendered slides')
@opt('--quality', default=DEFAULT_QUALITY, type=int, show_default=True,
     help='jpeg quality from 1 to 100 (ignored for png)')
@opt('-p', '--pages', default=None,
     help='pages to convert, 1-based, e.g. 1-5,8,12- (default: all)')
@opt('-s', '--size', 'slide_size', default='auto', show_default=True,
     help='slide aspect ratio: auto, or a ratio such as 16:9 or 4:3')
@opt('-n', '--notes', is_flag=True, default=False,
     help="copy each page's text into the slide's speaker notes")
@opt('--password', default=None,
     help='password for an encrypted pdf')
@opt('-y', '--overwrite', is_flag=True, default=False,
     help='overwrite the output file if it already exists')
@opt('-q', '--quiet', is_flag=True, default=False,
     help='disable printing progress bar and other info')
@opt('--from', 'start_page', default=None, type=int,
     help='first page in the pdf to copy, 0-based (deprecated: use --pages)')
@opt('--count', 'page_count', default=None, type=int,
     help='number of pages in the pdf to copy (deprecated: use --pages)')
@arg('pdf_files', nargs=-1, required=True,
     type=click.Path(exists=True, dir_okay=False))
def main(pdf_files, output_file, output_dir, resolution, image_format, quality,
         pages, slide_size, notes, password, overwrite, quiet, start_page,
         page_count):
    """
    Convert a PDF slideshow to Powerpoint PPTX.

    Renders each page as an image and creates the resulting Powerpoint
    slideshow from these images. Useful when you want to use Powerpoint
    to present a set of PDF slides (e.g. slides from Beamer). You can then
    use the presentation capabilities of Powerpoint (notes, ink on slides,
    etc.) with slides created in LaTeX.

    Accepts more than one PDF_FILE, in which case each is converted in turn.
    """
    if pages is not None and (start_page is not None or page_count is not None):
        raise click.UsageError(
            'use either --pages or --from/--count, not both')

    if output_file is not None and len(pdf_files) > 1:
        raise click.UsageError(
            '--output takes a single file: use --output-dir for several pdfs')

    if output_dir is not None:
        Path(output_dir).mkdir(parents=True, exist_ok=True)

    failures = 0
    for pdf_file in pdf_files:
        destination = _output_path(pdf_file, output_file, output_dir)
        if destination.exists() and not overwrite:
            print(f'{destination} already exists, use --overwrite to replace it',
                  file=sys.stderr)
            failures += 1
            continue
        try:
            convert_pdf2pptx(
                pdf_file, destination, resolution, start_page, page_count,
                quiet, pages=pages, image_format=image_format, quality=quality,
                notes=notes, password=password, slide_size=slide_size)
        except (ValueError, PermissionError, RuntimeError) as err:
            print(f'{pdf_file}: {err}', file=sys.stderr)
            failures += 1

    if failures:
        sys.exit(1)


if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter
