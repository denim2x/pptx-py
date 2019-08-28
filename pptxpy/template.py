# encoding: utf-8

from types import MethodType as bind
from pptx import Presentation as _load
from .common import Slides, RT, Cache, PresentationPart


class Template:
  def __init__(self, uri):
    self._uri = uri
    self._model = None

  def __call__(self):
    prs = _load(self._uri)
    if self._model is None:
      self._model = _Slides(prs.slides)

    prs.part.drop_all(RT.SLIDE)
    prs.slides.clear()
    prs.part._model = self._model

    return prs


def Slides_spawn(self, slide_index=None, slide_id=None):
  if self.part._model is None:
    return

  slide_model = self.part._model(slide_index, slide_id)
  if slide_model is None:
    return

  cache = Cache(self.part.package)
  slide_part = slide_model(self.part.package, cache)

  rId = self.part.relate_to(slide_part, RT.SLIDE)
  self._sldIdLst.add_sldId(rId)

  return slide_part.slide


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


class _Part(_Model):
  def __init__(self, part, owner=None, slide_id=None):
    part._model = self
    self._load = part.load
    self._partname = part.partname
    self._uri = self.partname.template
    self._content_type = part.content_type
    self._blob = part.blob
    self._owner = owner
    self._slide_id = slide_id
    self._rels = [_Relationship(rel, self) for rel in part.rels.values()]

  @property
  def slide_id(self):
    return self._slide_id

  @property
  def partname(self):
    return self._partname

  @property
  def content_type(self):
    return self._content_type

  def __call__(self, package, cache=None):
    if cache is not None:
      uri = cache.next_partname(self._uri)
    else:
      uri = package.next_partname(self._uri)
    part = self._load(uri, self._content_type, self._blob, package)

    for rel in self._rels:
      rel(part, cache)

    return part

  @classmethod
  def get(cls, part, owner):
    if hasattr(part, '_model'):
      return part._model

    return cls(part, owner)


class _Relationship(_Model):
  def __init__(self, rel, owner=None):
    self._rId = rel.rId
    self._reltype = rel.reltype
    self._is_external = rel.is_external
    if not self.is_external:
      self._target = _Part.get(rel.target_part, owner)
    else:
      self._target = rel.target_ref

  @property
  def is_external(self):
    return self._is_external

  @property
  def reltype(self):
    return self._reltype

  @property
  def target(self):
    return self._target

  def __call__(self, part, cache=None):
    target = self.target

    if not self.is_external:
      target = part.package[self.target, self.reltype]

      if target is None:
        target = cache[self.target] if cache else self.target(part.package)

    return part.rels.add_relationship(self._reltype, target, self._rId, self.is_external)


def _mount():
  Slides.__call__ = Slides_spawn
  PresentationPart._model = None
