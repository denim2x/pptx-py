# encoding: utf-8

"""Python library with various tools for enhancing python-pptx"""

__version__ = '0.0.1'

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

Rels._static = {
  RT.SLIDE, RT.IMAGE, RT.MEDIA, RT.VIDEO, RT.NOTES_MASTER#, RT.SLIDE_MASTER
}

Part._cached = {
  CT.PML_SLIDE_MASTER, CT.PML_SLIDE_LAYOUT
}

Rels._restricted = {
  RT.SLIDE_MASTER: { CT.PML_SLIDE_LAYOUT }
}

Part._closed = {
  CT.PML_SLIDE_MASTER: { RT.SLIDE_LAYOUT }
}

_void = set()

tmpl_re = re.compile(r"^(.+?)(\d+)?(\.\w+)?$")
name_re = re.compile(r"^(?:(\d+)_)?")
idLstItem_tag = NamespacePrefixedTag('p:sldLayoutId').clark_name


def Slides_duplicate(self, slide_index=None, slide_id=None, slide_master=False):
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
  prs = self.parent
  if not hasattr(prs.part, '_cache'):
    prs.part._cache = {}

  if not hasattr(prs.part, '_max_sldId'):
    max_sldId = 0
    if len(prs.slide_masters) > 0:
      master = prs.slide_masters[-1]
      layout_ids = master.slide_layouts._sldLayoutIdLst
      if len(layout_ids) > 0:
        max_sldId = int(layout_ids[-1].attrib['id'])
    prs.part._max_sldId = max_sldId
  cloner = Cloner(prs.part, slide_master)
  # cloner = Cloner(parts, part._cache, [prs.slide_masters[m].part if isinstance(m, int) else m for m in slide_masters] if slide_masters else None)
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
  if cloner is None:
    return self._clone(uri)

  if self not in cloner:
    part = self._clone(uri)
    cloner[part] = self
    return part

  return self

Part.clone = Part_clone


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


def Part__clone(self, uri=None):
  """
  Creates a _shallow_ duplicate of *self*, optionally having *partname* assigned
  the value of *uri* (if non-null), otherwise *self.partname*.

  Return value: The newly created |Part| instance.
  """
  if uri is None:
    uri = self.partname
  
  blob = self.blob  
  if self.content_type == CT.OFC_THEME:
    xml = parse_xml(self.blob)
    name = xml.attrib['name']
    xml.attrib['name'] = name_re.sub(lambda m: "%d_" % (int(m.group(1) or '0') + 1), name)
    blob = dump_xml(xml)
  
  return self.load(uri, self.content_type, blob, self.package)

Part._clone = Part__clone


def Rels_append(self, rel):
  """
  Creates a new |_Relationship| instance based on *rel* and inserts it into *self*.
  
  Return value: A Boolean value indicating whether *rel is None*.
  """
  if rel is None:
    return False

  self.add_relationship(rel.reltype, rel._target, rel.rId, rel.is_external)
  return True

Rels.append = Rels_append


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


class Cloner:
  """
  Utility class for handling the cloning process for a given |_Relationship| 
  instance; uses a *_cache* to store all cloned |Part| instances - thus 
  avoiding _infinite recursion_.
  """
  def __init__(self, prs, slide_master=False):
    self._idx = {}
    for part in prs.package.parts:
      uri = part.partname
      tmpl = uri.template
      self._idx[tmpl] = max(self._idx.get(tmpl, 0), uri.index)
    self._cache = set()
    self._gcache = prs._cache
    # self._slide_masters = set(slide_masters) if slide_masters is not None else None
    self._slide_master = slide_master
    self._prs = prs

  def __setitem__(self, dest, src):
    if src is None:
      return

    part = None
    #content_type = None
    if isinstance(src, Part):
      if src.content_type in Part._cached:
        self._gcache[src] = dest
      else:
        self._cache.add(src)
      part = src
      #content_type = src.content_type

    rels = self._get_rels(src)
    if isinstance(rels, dict):
      rels = rels.values()
    
    if rels is None:
      return

    dest = self._get_rels(dest)
    ct = part.content_type if part else None
    try:
      for rel in rels:
        if rel.reltype in Part._closed.get(ct, _void):
          continue
        dest.attach(self._clone(rel, part))

    except TypeError:
      pass

  @classmethod
  def _get_rels(cls, self):
    rels = self

    if isinstance(self, PartElementProxy):
      rels = self.part

    if isinstance(self, Part):
      rels = self.rels

    return rels

  def __contains__(self, part):
    return part in self._cache

  def _clone(self, rel, src):
    ct = src.content_type if src else None
    if rel.is_external:
      target = rel.target_ref
    else:
      target = rel.target_part
      if self._cloneable(rel, ct):
        # if target in self._slide_masters:
        #   target = self._clone_part(target)
        #   if ct == CT.PML_SLIDE_LAYOUT:
        #     target.rels.get_or_add(RT.SLIDE_LAYOUT, src)

        # elif not rel.is_static:
        # if not rel.is_static and (target.content_type != CT.PML_SLIDE_MASTER or self._slide_masters is None or target in self._slide_masters):
        if target in self._gcache:
          target = self._gcache[target]
        elif not rel.is_static and (target.content_type != CT.PML_SLIDE_MASTER or self._slide_master):
          target = self._clone_part(target)
          if target.content_type == CT.PML_SLIDE_MASTER:
            for item in target.slide_master.slide_layouts._sldLayoutIdLst.iterchildren():
              item.delete()

        if target.content_type == CT.PML_SLIDE_MASTER and ct == CT.PML_SLIDE_LAYOUT and src in self._gcache:
          r = target.rels.get_or_add(RT.SLIDE_LAYOUT, self._gcache[src])
          id_list = target.slide_master.slide_layouts._sldLayoutIdLst
          self._prs._max_sldId += 1
          item = id_list.makeelement(idLstItem_tag, { 'id': str(self._prs._max_sldId), qn('r:id'): r.rId }, id_list.nsmap)
          id_list.append(item)

    return Rel(rel.rId, rel.reltype, target, rel._baseURI, rel.is_external)

  def _clone_part(self, part):
    uri = part.partname
    tmpl = uri.template
    self._idx[tmpl] += 1
    uri = PackURI(tmpl % self._idx[tmpl])
    return part.clone(uri, self)

  @classmethod
  def _cloneable(cls, rel, content_type):
    #return True
    return content_type in Rels._restricted.get(rel.reltype, { content_type })


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
