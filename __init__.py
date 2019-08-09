from lxml import etree
from pptx.slide import Slides
from pptx.parts.slide import SlidePart
from pptx.opc.packuri import PackURI
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.presentation import PresentationPart  #FIXME: Attach *duplicate()* here also


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

  max_id = 0 if new_ids else None
  max_uri, max_idx = None, 0
  for slide in self:
    if new_ids:
      max_id = max(max_id, *iter_ids(slide))
    uri = partname(slide)
    if max_idx < uri.idx:
      max_uri, max_idx = uri, uri.idx

  slide_part = clone(slide._part, max_uri, max_id)

  part = self.part
  rId = part.relate_to(slide_part, RT.SLIDE)
  self._sldIdLst.add_sldId(rId)

  return slide_part.slide

Slides.duplicate = duplicate

def clone(self, base_uri=None, base_id=None):
  """
  Creates an exact copy of *self* (|SlidePart|) by building a new |SlidePart|
  instance from *self*, optionally increasing all ID values by *base_id*, if 
  specified. The new |SlidePart|'s *partname* is *base_uri* increased by 1,
  if specified.
  
  Return value: The newly created |SlidePart| instance.
  """

  uri = None
  if base_uri is not None:
    uri = PackURI('%s%d.%s' % (basename(base_uri), base_uri.idx + 1, base_uri.ext)
  part = SlidePart.load(uri, self.content_type, self.blob, self.package)

  if base_id is not None:
    pass   # FIXME: Finish implementation

  return part

def partname(self):
  return self._part.partname

def iter_ids(self):
  for e in xpath(self, '//@id'):
    yield int(e)

def xpath(self, expr):
  return self.element.xpath(expr)

def basename(self):
  filename = self.filename
  if not filename:
      return None
  return posixpath.splitext(filename)[0]
