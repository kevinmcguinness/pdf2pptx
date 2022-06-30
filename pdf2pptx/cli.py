#!/usr/bin/env python
# Copyright (c) 2020 Kevin McGuinness <kevin.mcguinness@gmail.com>

import click
import sys

from . import convert_pdf2pptx

DEFAULT_RESOLUTION = 300
DEFAULT_START_PAGE = 0
DEFAULT_QUIET = False
DEFAULT_PAGE_COUNT = None
DEFAULT_OUTPUT = None
arg = click.argument
opt = click.option


@click.command()
@opt('-o', '--output', 'output_file', default=DEFAULT_OUTPUT,
     help='location to save the pptx (default: PDF_FILE.pptx)')
@opt('-r', '--resolution', default=DEFAULT_RESOLUTION, type=int,
     help=f'resolution in dots per inch (default: {DEFAULT_RESOLUTION})')
@opt('-q', '--quiet', is_flag=True, default=DEFAULT_QUIET,
     help='disable printing progress bar and other info')
@opt('--from', 'start_page', default=DEFAULT_START_PAGE, type=int,
     help='first page in the pdf to copy to the pptx')
@opt('--count', 'page_count', default=DEFAULT_PAGE_COUNT, type=int,
     help='number of pages in the pdf to copy to the pptx')
@arg('pdf_file', type=click.Path(exists=True, dir_okay=False))
def main(pdf_file, output_file, resolution, start_page, page_count, quiet):
    """
    Convert a PDF slideshow to Powerpoint PPTX.

    Renders each page as a PNG image and creates the resulting Powerpoint 
    slideshow from these images. Useful when you want to use Powerpoint
    to present a set of PDF slides (e.g. slides from Beamer). You can then
    use the presentation capabilities of Powerpoint (notes, ink on slides,
    etc.) with slides created in LaTeX.
    """
    try:
        convert_pdf2pptx(
            pdf_file, output_file, resolution, start_page, page_count, quiet)
    except PermissionError as err:
        print(err, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter
