from setuptools import setup

setup(
   name='pdf2pptx',
   version='1.0',
   description='Utility to convert a PDF slideshow to Powerpoint PPTX.',
   author='Kevin McGuinness',
   author_email='kevin.mcguinness@gmail.com',
   packages=[], 
   install_requires=['pymupdf', 'python-pptx'],
)
