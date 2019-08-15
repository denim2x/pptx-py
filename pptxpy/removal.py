# encoding: utf-8

from .common import Slides


def Slides_remove(self, slide_index=None, slide_id=None, sweep=True):
  """
  """
  slide = self._get(slide_index, slide_id)
  if slide is None:
    return

  prs = self.parent

  dropped = prs.part.drop(slide.part)
  sldIds = set()
  for item in self._sldIdLst:
    if item.rId in dropped:
      sldIds.add(item)

  for item in sldIds:
    item.delete()

  if sweep:
    for s in self:
      s.part.drop(slide.part)

  return slide

def _mount():
  Slides.remove = Slides_remove
