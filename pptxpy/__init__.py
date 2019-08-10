# encoding: utf-8

"""Initialization module for python-pptx package."""

__version__ = '0.0.1'

try:
  import pptx
except ImportError:
  raise Exception("Module pptx-py requires python-pptx in order to run; please install it first, then try again")

import posixpath

from pptx.slide import Slide, Slides
from pptx.parts.slide import SlidePart
from pptx.opc.packuri import PackURI
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.presentation import PresentationPart  #FIXME: Attach *duplicate()* here also
from pptx.opc.package import _Relationship as Rel, RelationshipCollection as Rels


def duplicate(self, slide_index=None, slide_id=None, new_ids=False):
  """
  Creates an _identical_ copy of the |Slide| at *slide_index* (or *slide_id*) 
  by cloning its corresponding |SlidePart| instance, then inserts it into *self*.
  Optionally creates new element IDs, according to *new_ids*.

  Return value: the newly created |Slide| instance.
  """
  slide = None

  if slide_index is not None:
      slide = self[slide_index]
  elif slide_id is not None:
      slide = self.get(slide_id)

  if slide is None:
      return  

  max_id = None
  if new_ids:
    max_id = 0
    for slide in self:
      max_id = max(max_id, *iter_ids(slide))
  
  part = self.part
  slide_part = clone(slide.part, part._next_slide_partname, max_id)

  rId = part.relate_to(slide_part, RT.SLIDE)
  self._sldIdLst.add_sldId(rId)

  return slide_part.slide

Slides.duplicate = duplicate

def clone(self, uri=None, base_id=None):
  """
  Creates an exact copy of *self* (|SlidePart|) by building a new |SlidePart|
  instance from *self*, optionally increasing all ID values by *base_id*, if 
  specified. The new |SlidePart|'s *partname* is *base_uri* increased by 1,
  if specified.
  
  Return value: The newly created |SlidePart| instance.
  """

  part = SlidePart.load(uri, self.content_type, self.blob, self.package)
  part.rels.assign(self)

  if base_id is not None:
    pass   # FIXME: Finish implementation

  return part

def Rels_assign(self, src):
  """
  Assigns all |_Relationship| instances from *src* to *self*
  """
  if src is None:
    return self

  if isinstance(src, Slide):
    src = src.part

  if isinstance(src, SlidePart):
    src = src.rels

  if isinstance(src, dict):
    for rId, rel in src.items():
      self.add_relationship(rel.reltype, rel._target, rId, rel.is_external)
  else:
    for rel in src:
      self.add_relationship(rel.reltype, rel._target, rel.rId, rel.is_external)

  return self

Rels.assign = Rels_assign

def iter_ids(self):
  for e in xpath(self, '//@id'):
    yield int(e)

def xpath(self, expr):
  return self.element.xpath(expr)

def Rels_eq(self, other):
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, Rels):
    return False

  if len(self) != len(other):
    return False

  for rId, rel in self.items():
    if not rId in other:
      return False
    if rel != other[rId]:
      return False

  return True

Rels.__eq__ = Rels_eq

def Rel_eq(self, other):
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, Rel):
    return False

  if self.reltype != other.reltype:
    return False

  if self._target != other._target:
    return False

  if self.is_external != other.is_external:
    return False

  return True

Rel.__eq__ = Rel_eq
