# pdf2pptx

Utility to convert a PDF slideshow to Powerpoint PPTX.

Renders each page as a PNG image and creates the resulting Powerpoint 
slideshow from these images. Useful when you want to use Powerpoint 
to present a set of PDF slides (e.g. slides from Beamer). You can then
use the presentation capabilities of Powerpoint (notes, ink on slides,
etc.) with slides created in LaTeX.

I use this to present PDF slides on a Surface Pro so that I can annotate
them with a stylus as I present.

Uses [PyMuPDF](https://github.com/pymupdf/PyMuPDF) and 
[python-pptx](https://github.com/scanny/python-pptx) to do the hard work.

