# encoding: utf-8

try:
  import pptx
except ImportError:
  raise Exception("Module pptx-py requires python-pptx in order to run. Install it first, then try again.")

import posixpath, re

from pptx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
from pptx.opc.oxml import serialize_part_xml as dump_xml
from pptx.opc.package import _Relationship as Rel, RelationshipCollection as Rels, Part
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.shared import PartElementProxy
from pptx.slide import Slide, Slides
from pptx.oxml.ns import NamespacePrefixedTag, qn


_void = set()

tmpl_re = re.compile(r"^(.+?)(\d+)?(\.\w+)?$")
name_re = re.compile(r"^(?:(\d+)_)?")
idLstItem_tag = NamespacePrefixedTag('p:sldLayoutId').clark_name


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

Part.is_similar = Part_is_similar


@property
def Part_basename(self):
  return posixpath.basename(self.partname)

Part.basename = Part_basename


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

Rels.attach = Rels_attach


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

Rels.__eq__ = Rels_eq
Rels.equals = Rels_eq


def Rels_pprint(self):
  return '%s{\n  %s\n}' % (Rels.__name__, '\n'.join('%s: %s' % (rId, rel.pprint()) for rId, rel in self.items()))

Rels.pprint = Rels_pprint

@property
def Rel_is_static(self):
  return self.reltype in Rels._static

Rel.is_static = Rel_is_static


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

Rel.__eq__ = Rel_eq
Rel.equals = Rel_eq


def Rel_pprint(self):
  reltype = posixpath.basename(self.reltype)
  target = self.target.target_ref if self.is_external else self.target_part.partname
  return '%s{ reltype="…/%s" target="%s" baseURI="%s" is_external=%s }' % (
    Rel.__name__, reltype, target, self._baseURI, self.is_external
  )

Rel.pprint = Rel_pprint


@property
def PackURI_index(self):
  if not hasattr(self, '_index'):
    self._index = int(tmpl_re.match(self).group(2) or '0')

  return self._index

PackURI.index = PackURI_index


@property
def PackURI_template(self):
  return tmpl_re.sub(r'\1%d\3', self)

PackURI.template = PackURI_template


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

PackURI.is_similar = PackURI_is_similar


def Slide_is_similar(self, other):
  if self is None:
    return other is None

  if other is None:
    return False

  if not isinstance(other, Slide):
    return False

  return self.part.is_similar(other.part)

Slide.is_similar = Slide_is_similar
