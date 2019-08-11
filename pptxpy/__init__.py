# encoding: utf-8

"""Python library with various tools for enhancing python-pptx"""

__version__ = '0.0.1'

try:
  import pptx
except ImportError:
  raise Exception("Module pptx-py requires python-pptx in order to run. Install it first, then try again.")

import posixpath, re

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import _Relationship as Rel, RelationshipCollection as Rels, Part
from pptx.opc.packuri import PackURI
from pptx.shared import PartElementProxy
from pptx.slide import Slides

static_rels = {
  RT.IMAGE, RT.MEDIA, RT.VIDEO, RT.SLIDE_LAYOUT, RT.NOTES_MASTER, RT.SLIDE_MASTER
}


def Slides_duplicate(self, slide_index=None, slide_id=None):
  """
  Creates an _identical_ copy of the |Slide| instance (given by either *slide_index*
  _or_ *slide_id*) by cloning its corresponding |SlidePart| instance, then appends
  it to *self*.

  Return value: the newly created |Slide| instance.
  """
  slide = None

  if slide_index is not None:
      slide = self[slide_index]
  elif slide_id is not None:
      slide = self.get(slide_id)

  if slide is None:
      return 
  
  part = self.part
  parts = part.package.parts
  cloner = Cloner(parts)
  slide_part = slide.part.clone(part._next_slide_partname, cloner)

  rId = part.relate_to(slide_part, RT.SLIDE)
  self._sldIdLst.add_sldId(rId)

  return slide_part.slide

Slides.duplicate = Slides_duplicate


def Part_clone(self, uri=None, cloner=None):
  """
  Creates an exact copy of this |Part| instance. The *partname* of the new instance
  is *uri* if non-null, otherwise *self.partname*.
  
  Return value: The newly created |Part| instance.
  """
  if uri is None:
    uri = self.partname

  if self not in cloner:
    part = self.load(uri, self.content_type, self.blob, self.package)
    part.rels.assign(self, cloner + self)
    return part

  return self

Part.clone = Part_clone


def Rels_assign(self, src, cloner=None):
  """
  Assigns all |_Relationship| instances from *src* to *self*; optionally
  creates clones of all non-static target parts (when *cloner* is non-null)
  """
  if src is None:
    return self

  if isinstance(src, PartElementProxy):
    src = src.part

  if isinstance(src, Part):
    src = src.rels

  if isinstance(src, dict):
    src = src.values()

  try:
    for rel in src:
      if cloner:
        self.attach(cloner(rel))
      else:
        self.append(rel)

  except TypeError:
    raise

  return self

Rels.assign = Rels_assign


class Cloner:
  def __init__(self, parts):
    self._parts = parts
    self._idx = {}
    self._cache = set()

  def __contains__(self, part):
    return part in self._cache

  def __add__(self, part):
    self._cache.add(part)
    return self

  def __call__(self, rel):
    target = rel.target_part
    uri = target.partname
    if not rel.is_static and uri.idx is not None:
      tmpl = re.sub(r'^(.+?)(\d+)(\.\w+)$', r'\1%d\3', uri)
      if tmpl not in self._idx:
        max_idx = 0
        for part in self._parts:
          if part.matches(tmpl, max_idx):
            max_idx = part.partname.idx
        self._idx[tmpl] = max_idx

      self._idx[tmpl] += 1
      uri = PackURI(tmpl % self._idx[tmpl])
      target = target.clone(uri, self)

    return Rel(rel.rId, rel.reltype, target, rel._baseURI, rel.is_external)


def Part_matches(self, tmpl, max_idx=None):
  uri = self.partname
  idx = uri.idx

  if max_idx is not None:
    return False if idx is None else uri == tmpl % idx and max_idx < idx

  return uri == tmpl if idx is None else uri == tmpl % idx

Part.matches = Part_matches


def Part_is_similar(self, other):
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, Part):
    return False

  if self.partname.baseURI != other.partname.baseURI:
    return False

  if self.content_type != other.content_type:
    return False

  if self.blob != other.blob:
    return False

  if self.package != other.package:
    return False

  return True

Part.is_similar = Part_is_similar


def Rels_append(self, rel):
  if rel is None:
    return False

  self.add_relationship(rel.reltype, rel._target, rel.rId, rel.is_external)
  return True

Rels.append = Rels_append


def Rels_attach(self, rel):
  target = rel.target_part

  self[rel.rId] = rel
  if not rel.is_external:
    self._target_parts_by_rId[rel.rId] = target

  return target

Rels.attach = Rels_attach


@property
def Rel_is_static(self):
  return self.reltype in static_rels

Rel.is_static = Rel_is_static


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

  if not self.target_part.is_similar(other.target_part):
    return False

  if self.is_external != other.is_external:
    return False

  return True

Rel.__eq__ = Rel_eq
