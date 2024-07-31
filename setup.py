#!/usr/bin/env python
# -*- coding: utf-8 -*-
import io
import os
import sys

from setuptools import setup

here = os.path.abspath(os.path.dirname(__file__))
with io.open(os.path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = '\n' + f.read()

setup(
   name='pdf2pptx',
   version='1.0.5',
   description='Utility to convert a PDF slideshow to Powerpoint PPTX.',
   long_description=long_description,
   long_description_content_type='text/markdown',
   author='Kevin McGuinness',
   url='https://github.com/kevinmcguinness/pdf2pptx',
   license='MIT',
   author_email='kevin.mcguinness@gmail.com',
   packages=['pdf2pptx'],
   install_requires=['pymupdf', 'python-pptx', 'click', 'tqdm'],
   entry_points={
       'console_scripts': ['pdf2pptx=pdf2pptx.cli:main'],
   },
   classifiers=[
        # Trove classifiers
        # Full list: https://pypi.python.org/pypi?%3Aaction=list_classifiers
        'Development Status :: 4 - Beta',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Environment :: Console',
        'License :: OSI Approved :: MIT License',
        'Topic :: Utilities',
        'Topic :: Multimedia :: Graphics :: Presentation',
    ],
)
