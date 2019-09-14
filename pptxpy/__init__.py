# encoding: utf-8

"""Python library with various tools for enhancing python-pptx"""

__version__ = '0.0.1'

from .common import _mount
_mount()

from .cloning import _mount
_mount()

from .removal import _mount
_mount()

from .template import Template

del _mount
