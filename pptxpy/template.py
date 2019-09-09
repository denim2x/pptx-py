# encoding: utf-8

from posixpath import splitext
from pptx import Presentation as _load
from .common import Slides, RT, Cache, Presentation, PresentationPart, PackURI, Slide, Part, Rel
from .slide import _Slide


class Template:
  def __init__(self, uri):
    self._uri = uri
    self._model = None

  def __call__(self):
    prs = _load(self._uri)     # FIXME: Check _uri existence
    if self._model is None:
      self._model = _Slides(prs.slides)

    for node in prs.slide_layouts:
      node.part.drop_all(RT.SLIDE, recursive=False)

    prs.part.drop_all(RT.SLIDE)
    prs.slides.clear()
    prs.part._model = self._model

    return _Presentation(prs.part)


def Slides_spawn(self, slide_index=None, slide_id=None):
  if self.part._model is None:
    return

  slide_model = self.part._model(slide_index, slide_id)
  if slide_model is None:
    return

  slide_part = slide_model(Cache(self.part.package))

  rId = self.part.relate_to(slide_part, RT.SLIDE)
  self._sldIdLst.add_sldId(rId)

  return _Slide(slide_part)
  # return slide_part.slide


class _Presentation(Presentation):
  def __init__(self, part):
    Presentation.__init__(self, part._element, part)
    part._presentation = self

  def save(self, file, update_links=False):  # FIXME: Clean extraneous relationships (including those without corresponding links)
    for slide in self.slides:
      slide._relink(self.slides, update_links)

    self.part.save(file)
    return self

class _Model:
  pass

class _Slides(_Model):
  def __init__(self, slides):
    self._list = [_Part(s.part, self, s.slide_id) for s in slides]
    self._ids = {}

  def get(self, slide_id):
    if slide_id not in self._ids:
      for item in self._list:
        if item.slide_id == slide_id:
          self._ids[slide_id] = item
          break

    return self._ids[slide_id]

  def __getitem__(self, slide_index):
    return self._list[slide_index]

  def __iter__(self):
    return self._list.values()

  def __len__(self):
    return len(self._list)

  def __call__(self, slide_index=None, slide_id=None):
    """
    Create a new |Slide| instance from the Slide model given by either
    (but not both) *slide_index* or *slide_id*.
    """
    model = None
    if slide_index is not None:
      model = self[slide_index]
    elif slide_id is not None:
      model = self.get(slide_id)

    return model


class _Reference(_Model):
  def __init__(self, part, owner=None, slide_id=None):
    base, ext = splitext(part.partname)
    self._partname = PackURI('%s%s' % (base, ext.lower()) if ext else base)
    self._content_type = part.content_type
    self._blob = part.blob
    self._owner = owner
    self._slide_id = slide_id

  @property
  def blob(self):
    return self._blob

  @property
  def slide_id(self):
    return self._slide_id

  @property
  def partname(self):
    return self._partname

  @property
  def content_type(self):
    return self._content_type


class _Part(_Reference):
  def __init__(self, part, owner=None, slide_id=None):
    _Reference.__init__(self, part, owner, slide_id)
    part._model = self
    self._uri = self.partname.template
    self._load = part.load    
    self._rels = [_Relationship(rel, self) for rel in part.rels.values()]

  def __call__(self, cache):
    uri = cache.next_partname(self._uri)
    part = self._load(uri, self._content_type, self._blob, cache.package)

    cache[self] = part

    for rel in self._rels:
      if rel.reltype == RT.SLIDE:
        continue

      out = rel(part, cache)
      if rel.reltype in part._reltypes:
        out.target_part.relate_to(part, part._reltype)

    return part

  @classmethod
  def get(cls, part, owner):
    if hasattr(part, '_model'):
      return part._model

    return cls(part, owner)


class _Relationship:
  def __init__(self, rel, owner=None):
    self._reltype = rel.reltype
    self._is_external = rel.is_external
    if self.reltype != RT.SLIDE or isinstance(owner, _Model):
      self._rId = rel.rId
      if self.reltype == RT.SLIDE:
        self._target = _Reference(rel.target_part, owner)
      elif not self.is_external:
        self._target = _Part.get(rel.target_part, owner)
      else:
        self._target = rel.target_ref
    else:
      self._rId = None
      self._target = owner

  @property
  def is_external(self):
    return self._is_external

  @property
  def reltype(self):
    return self._reltype

  @property
  def target(self):
    return self._target  

  def __getitem__(self, target):
    if target:
      return _Relationship(self, target)

  def __call__(self, part, cache=None):
    target = self.target
    if not self.is_external and self.reltype != RT.SLIDE:
      target = part.package[self.target, self.reltype]

      if target is None and cache is not None:
        target = cache[self.target]

    return part.load_rel(self.reltype, target, self._rId, self.is_external)


def _mount():
  Slides.__call__ = Slides_spawn
  PresentationPart._model = None
