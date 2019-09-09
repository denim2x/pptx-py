# encoding: utf-8

from .common import Slide

class _Slide(Slide):
  def __init__(self, part):
    Slide.__init__(self, part._element, part)
    self._links = {}

  def __getitem__(self, link_id):
    return self._links[link_id]

  def __setitem__(self, link_id, slide_id):
    self._links[link_id] = slide_id

  def __delitem__(self, link_id):
    del self._links[link_id]

  def __contains__(self, link_id):
    return link_id in self._links

  def __iter__(self):
    return iter(self._links)

  def __len__(self):
    return len(self._links)

  def relink(self, links_dict=None, **links):
    if links_dict:
      self._links.update(links_dict)
    self._links.update(links)
    return self
