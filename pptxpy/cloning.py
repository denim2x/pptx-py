# encoding: utf-8

from .common import Part, Rels, Rel, PartElementProxy, Slides, CT, RT, PackURI
from .common import qn, _void, parse_xml, name_re, dump_xml, idLstItem_tag


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

  cloner = Cloner(prs.part, slide_master)
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
    # self._slide_masters = set(slide_masters) if slide_masters is not None else None
    self._slide_master = slide_master

    if not hasattr(prs, '_cache'):
      prs._cache = {}
    self._gcache = prs._cache

    if not hasattr(prs, '_related'):
      prs._related = {}
    self._rels = prs._related

    _prs = prs.presentation
    if not hasattr(prs, '_max_sldId'):
      max_sldId = 0
      if len(_prs.slide_masters) > 0:
        master = _prs.slide_masters[-1]
        layout_ids = master.slide_layouts._sldLayoutIdLst
        if len(layout_ids) > 0:
          max_sldId = int(layout_ids[-1].attrib['id'])
      prs._max_sldId = max_sldId

    self._prs = prs

  def __setitem__(self, dest, src):
    if src is None:
      return

    part = None
    if isinstance(src, Part):
      if src.content_type in Part._cached:
        self._gcache[src] = dest
      else:
        self._cache.add(src)
      part = src

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
        if target in self._gcache:
          target = self._gcache[target]
        elif not rel.is_static and (target.content_type != CT.PML_SLIDE_MASTER or self._slide_master):
          target = self._clone_part(target)
          if target.content_type == CT.PML_SLIDE_MASTER:
            for item in target.slide_master.slide_layouts._sldLayoutIdLst.iterchildren():
              item.delete()

        if target.content_type == CT.PML_SLIDE_MASTER and ct == CT.PML_SLIDE_LAYOUT and src in self._gcache:
          rId = target.relate_to(self._gcache[src], RT.SLIDE_LAYOUT)
          r = target.rels[rId]
          id_list = target.slide_master.slide_layouts._sldLayoutIdLst
          self._prs._max_sldId += 1
          item = id_list.makeelement(idLstItem_tag, { 'id': str(self._prs._max_sldId), qn('r:id'): r.rId }, id_list.nsmap)
          id_list.append(item)

        if target not in self._rels:
          self._rels[target] = self._prs.relate_to(target, rel.reltype)
      
    return Rel(rel.rId, rel.reltype, target, rel._baseURI, rel.is_external)

  def _clone_part(self, part):
    uri = part.partname
    tmpl = uri.template
    self._idx[tmpl] += 1
    uri = PackURI(tmpl % self._idx[tmpl])
    return part.clone(uri, self)

  @classmethod
  def _cloneable(cls, rel, content_type):
    return content_type in Rels._restricted.get(rel.reltype, { content_type })

