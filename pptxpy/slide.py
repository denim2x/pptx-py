# encoding: utf-8

from .common import Slide, RT

_hlinksldjump = "//*[@action='ppaction://hlinksldjump']"


class _Slide(Slide):
  def __init__(self, part, sldId):
    Slide.__init__(self, part._element, part)
    part._slide = self
    self._links = {}
    self._sldId = sldId
    for link in self.slide_jumps:
      link.rId = None  

  def relocate(self, position=None):
    sldId = self._sldId
    sldId_list = sldId.getparent()
    if sldId_list is None:
      return False

    count = len(sldId_list)
    append = False
    if position is None:
      append = True
      position = -1
    
    if position < 0:
      position += count

    pos = sldId_list.index(sldId)
    if pos == position:
      return False

    del sldId_list[pos]
    if append:
      sldId_list.append(sldId)
    else:
      sldId_list.insert(position, sldId)

    return True

  def remove(self, position=None):
    sldId = self._sldId
    sldId_list = sldId.getparent()
    if sldId_list is None:
      return False

    sldId_list.remove(sldId)
    self.owner_part.drop(self.part)
    return True

  @property
  def owner_part(self):
    return self.part.package.presentation_part

  def __getitem__(self, link_id):
    return self._links[self._resolve(link_id)]

  def __setitem__(self, link_id, slide_id):
    key = self._resolve(link_id)
    if key is not None:
      self._links[key] = slide_id

  def __delitem__(self, link_id):
    return self._pop(link_id)

  def __contains__(self, link_id):
    return self._resolve(link_id) in self._links

  def __iter__(self):
    return iter(self._links.items())

  def __len__(self):
    return len(self._links)

  @property
  def slide_jumps(self):
    return self._xpath(_hlinksldjump)  

  def relink(self, links_dict=None, **links):
    self._update(links_dict)
    self._update(links)
    return self

  def unlink(self, id_list=None, *ids):
    if id_list is None and len(ids) == 0:
      self._links.clear()
      return self

    a = self._purge(id_list)
    b = self._purge(ids)

    return a or b

  def _xpath(self, sel):
    return self.part._element.xpath(sel)

  def _resolve(self, link_id):
    if not isinstance(link_id, int):
      return link_id

    links = self.slide_jumps
    if links:
      try:
        return links[link_id]
      except (TypeError, IndexError):
        pass

  def _relate(self, link, target, update):
    rId = link.rId
    if rId is None:
      rId = self.part.relate_to(target, RT.SLIDE)
      link.rId = rId
      return rId

    rels = self.part.rels
    if not update:
      if rId in rels:    # FIXME: Validate relationship
        return rId

      link.rId = None
      return

    if rId in rels:
      rels[rId]._target = target
    else:
      self.part.load_rel(RT.SLIDE, target, rId)

    return rId

  def _relink(self, slides, update=False):
    for link in self.slide_jumps:
      if link not in self._links:
        self._strip(link)
        link.rId = None
        continue

      slide_id = self._links[link]
      if isinstance(slide_id, int):
        target = slides.at(slide_id)
      else:
        target = slides.at(slide_id=slide_id)

      if target is None:
        if self._strip(link) and update:
          link.rId = None
        continue

      self._relate(link, target.part, update)

    return self

  def _strip(self, link):
    rId = link.rId
    rels = self.part.rels
    if rId in rels:
      del rels[rId]
      return True
      # link.rId = None
    return False

  def _pop(self, link_id):
    if link_id in self._links:
      del self._links[self._resolve(link_id)]
      return True

    return False

  def _purge(self, id_list):
    if id_list is None:
      return False

    ret = False
    for link_id in id_list:
      ret = ret or self._pop(link_id)

    return ret

  def _update(self, links):
    if links is None:
      return False

    for link_id, target in links.items():
      self[link_id] = target

    return True
