#!/usr/bin/env python

import os
import re
import glob

from setuptools import find_packages, setup


def ascii_bytes_from(path, *paths):
  """
  Return the ASCII characters in the file specified by *path* and *paths*.
  The file path is determined by concatenating *path* and any members of
  *paths* with a directory separator in between.
  """
  file_paths = glob.glob(os.path.join(path, *paths))
  if len(file_paths) > 1:
    return
  with open(file_paths[0]) as f:
    ascii_bytes = f.read()
  return ascii_bytes


# read required text from files
thisdir = os.path.dirname(__file__)
init_py = ascii_bytes_from(thisdir, "pptxpy", "__init__.py")
readme = ascii_bytes_from(thisdir, "README.*")
license = ascii_bytes_from(thisdir, "LICENSE")

# Read the version from pptx.__version__ without importing the package
# (and thus attempting to import packages it depends on that may not be
# installed yet)
version = re.search(r'''__version__ = ["']([^"']+)["']''', init_py).group(1)


NAME = "pptx-py"
VERSION = version
DESCRIPTION = "A set of useful tools for enhancing python-pptx"
KEYWORDS = "powerpoint ppt pptx office-tools openxml oxml python-pptx"
AUTHOR = "denim2x"
AUTHOR_EMAIL = "denim2x@cyberdude.com"
URL = "http://github.com/denim2x/pptx-py"
LICENSE = license
PACKAGES = find_packages(exclude=["tests", "tests.*"])

TEST_SUITE = "tests"
TESTS_REQUIRE = ["behave", "mock", "pyparsing>=2.0.1", "pytest"]

CLASSIFIERS = [
    "Development Status :: 2 - Pre-Alpha",
    "Environment :: Console",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: MIT License",
    "Operating System :: OS Independent",
    "Programming Language :: Python",
    "Programming Language :: Python :: 2",
    "Programming Language :: Python :: 2.7",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.6",
    "Topic :: Office/Business :: Office Suites",
    "Topic :: Software Development :: Libraries",
]

LONG_DESCRIPTION = readme + "\n\n"


params = {
    "name": NAME,
    "version": VERSION,
    "description": DESCRIPTION,
    "keywords": KEYWORDS,
    "long_description": LONG_DESCRIPTION,
    "author": AUTHOR,
    "author_email": AUTHOR_EMAIL,
    "url": URL,
    "license": LICENSE,
    "packages": PACKAGES,
    "tests_require": TESTS_REQUIRE,
    "test_suite": TEST_SUITE,
    "classifiers": CLASSIFIERS,
}

setup(**params)
