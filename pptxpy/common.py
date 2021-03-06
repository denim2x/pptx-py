# encoding: utf-8

try:
  import pptx
except ImportError:
  raise Exception("Module pptx-py requires python-pptx in order to run. Install it first, then try again.")

import posixpath, re
from collections import defaultdict

from pptx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
from pptx.opc.oxml import serialize_part_xml as dump_xml
from pptx.opc.package import _Relationship as Rel, RelationshipCollection as Rels, Part, OpcPackage
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.shared import PartElementProxy
from pptx.slide import Slide, Slides
from pptx.oxml.ns import NamespacePrefixedTag, qn
from pptx.presentation import Presentation
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import SlidePart
from pptx.util import lazyproperty

_void = set()

tmpl_re = re.compile(r"^(.+?)(\d+)?(\.\w+)?$")
name_re = re.compile(r"^(?:(\d+)_)?")

_media = {
  RT.IMAGE, RT.MEDIA, RT.VIDEO
}

_static = {
  RT.SLIDE_LAYOUT, RT.NOTES_MASTER, RT.SLIDE_MASTER, RT.CUSTOM_XML
} | _media


class Cache:
  def __init__(self, package):
    self._package = package
    self._partnames = defaultdict(lambda: 0)
    self._usednames = { p.partname for p in self.package.iter_parts() }
    self._parts = {}

  def next_partname(self, tmpl):
    self._partnames[tmpl] += 1
    while not self._available(tmpl):
      self._partnames[tmpl] += 1

    return self._uri(tmpl)

  def _available(self, tmpl):
    return self._uri(tmpl) not in self._usednames

  def _uri(self, tmpl):
    return PackURI(tmpl % self._partnames[tmpl])

  @property
  def package(self):
    return self._package

  def __getitem__(self, model):
    if model not in self._parts:
      model(self)
    return self._parts[model]

  def __setitem__(self, model, part):
    self._parts[model] = part


def Slides_at(self, slide_index=None, slide_id=None):
  """
  """
  slide = None

  if slide_index is not None:
    if isinstance(slide_index, int):
      if slide_index < 0:
        slide_index += len(self)
      if 0 <= slide_index < len(self):
        slide = self[slide_index]
  elif slide_id is not None:
    slide = self.get(int(slide_id))

  return slide

def Slides_clear(self):
  """
  Removes all slides from *self*.
  """
  return self._sldIdLst.clear()

def Slide_is_similar(self, other):
  if self is None:
    return other is None

  if other is None:
    return False

  return self.part.is_similar(other.part)


def Rels_attach(self, rel):
  """
  Inserts *rel* into *self*, performing additional necessary bindings.

  Return value: *rel.target_part* (or *rel.target_ref*, if *rel.is_external*).
  """
  self[rel.rId] = rel
  if rel.is_external:
    return rel.target_ref

  target = rel.target_part
  self._target_parts_by_rId[rel.rId] = target
  return target


def Rels_eq(self, other, rels=True):
  """
  Performs structural equality testing between *self* and *other*.

  Return value: The Boolean result of the tests.
  """
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, dict):
    return False

  if len(self) != len(other):
    return False

  for rId, rel in self.items():
    if not rId in other:
      return False
    #if rel != other[rId]:
    if not rel.equals(other[rId], rels):
      return False

  return True


def Rels_pprint(self):
  return '%s{\n  %s\n}' % (Rels.__name__, '\n'.join('%s: %s' % (rId, rel.pprint()) for rId, rel in self.items()))


@property
def Rel_is_static(self):
  return self.reltype in Rels._static


def Rel_eq(self, other, rels=True):
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, Rel):
    return False

  if self.reltype != other.reltype:
    return False

  if not self.target_part.is_similar(other.target_part, rels):
    return False

  if self.is_external != other.is_external:
    return False

  return True


def Rel_pprint(self):
  reltype = posixpath.basename(self.reltype)
  target = self.target.target_ref if self.is_external else self.target_part.partname
  return '%s{ reltype="…/%s" target="%s" baseURI="%s" is_external=%s }' % (
    Rel.__name__, reltype, target, self._baseURI, self.is_external
  )


@property
def PackURI_index(self):
  if not hasattr(self, '_index'):
    self._index = int(tmpl_re.match(self).group(2) or '0')

  return self._index


@property
def PackURI_template(self):
  return tmpl_re.sub(r'\1%d\3', self)


def PackURI_is_similar(self, other):
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, str):
    return False

  if not isinstance(other, PackURI):
    other = PackURI(other)

  return self.template == other.template

def OpcPackage_getitem(self, cursor):
  part, reltype = cursor
  uri = part.partname
  ct = part.content_type
  if reltype not in _static:
    return
  for part in self.iter_parts():
    if part.partname == uri and part.content_type == ct:
      return part

def Part_drop(self, part, exclude=None):
  dropped = set()
  for rel in self.rels.values():
    if exclude is not None and rel.reltype in exclude:
      continue

    if not rel.is_external and rel.target_part is part:
      dropped.add(rel)

  for rel in dropped:
    self.drop_rel(rel.rId)
    # del self.rels[rId]

  return dropped

def Part_drop_all(self, reltype, recursive=True):
  dropped = set()
  for rel in self.rels.values():
    if rel.reltype == reltype:
      dropped.add(rel)

  exclude = None if isinstance(recursive, bool) else recursive
  for rel in dropped:
    if recursive and not rel.is_external:
      for part in rel.target_part.related_parts:
        dropped.update(self.drop(part, exclude))

    self.drop_rel(rel.rId)
    # del self.rels[rId]

  return dropped

def Part_is_similar(self, other, rels=True):
  """
  Essentially performs shallow structural equality testing between
  *self* and *other* - with the exception of *partname* which is
  tested for _similarity_ rather then _equality_.

  Return value: The Boolean result of the tests.
  """
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, Part):
    return False

  if self.partname.is_similar(other.partname):
    return False

  if self.content_type != other.content_type:
    return False

  if self.blob != other.blob:
    return False

  if self.package != other.package:
    return False

  if rels and self.rels != other.rels:
    return False

  return True

@property
def Part_basename(self):
  return posixpath.basename(self.partname)


def _mount():
  Slides.at = Slides_at
  Slides.clear = Slides_clear
  Slide.is_similar = Slide_is_similar

  Rels.attach = Rels_attach
  Rels.__eq__ = Rels_eq
  Rels.equals = Rels_eq
  Rels.pprint = Rels_pprint

  Rel.is_static = Rel_is_static
  Rel.__eq__ = Rel_eq
  Rel.equals = Rel_eq
  Rel.pprint = Rel_pprint

  PackURI.index = PackURI_index
  PackURI.template = PackURI_template
  PackURI.is_similar = PackURI_is_similar

  OpcPackage.__getitem__ = OpcPackage_getitem

  SlidePart._reltypes = {
    RT.NOTES_SLIDE, #RT.SLIDE_LAYOUT
  }
  SlidePart._reltype = RT.SLIDE

  Part.drop = Part_drop
  Part.drop_all = Part_drop_all
  Part.is_similar = Part_is_similar
  Part.basename = Part_basename
  Part._reltypes = {}
  Part._reltype = None
