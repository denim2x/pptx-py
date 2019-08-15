# encoding: utf-8

from .common import Slides


def Slides_remove(self, slide_index=None, slide_id=None, erase=False):
  """
  """
  slide = self._get(slide_index, slide_id)
  if slide is None:
    return

  part = slide.part
  prs = self.parent

  dropped = set()
  for rel in prs.part.rels.values():
    if not rel.is_external and rel.target_part is part:
      dropped.add(rel.rId)

  _dropped = set()
  for item in self._sldIdLst:
    if item.rId in dropped:
      _dropped.add(item)

  for rId in dropped:
    del prs.part.rels[rId]

  for item in _dropped:
    item.delete()

  return slide

Slides.remove = Slides_remove
