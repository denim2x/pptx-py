import os
from pptx import Presentation

dirname = os.path.dirname(__file__)
try:
  import pptxy
except ImportError:
  import sys
  sys.path.insert(0, os.path.normpath(os.path.join(dirname, '..')))
  import pptxpy

prs = None

def setup(path):
  global prs
  prs = Presentation(normpath(path))
  num = len(prs.slides)
  for i in range(num):
    test_duplicate(i, True)

  import sys
  if len(sys.argv) > 1:
    path = sys.argv[1]
    prs.save(path)

def test_duplicate(i, muted=False):
  global prs
  s, num = prs.slides[i], len(prs.slides)
  c = prs.slides.duplicate(i)

  if muted:
    pass#return

  assert len(prs.slides) == num + 1
  assert prs.slides[-1] is c

  sp, cp = s.part, c.part
  assert sp.partname.is_similar(cp.partname)
  assert sp.content_type == cp.content_type
  assert sp.blob == cp.blob
  assert sp.package == cp.package
  assert sp.rels.equals(cp.rels, False)

  return
  assert sp.rels == cp.rels, \
    'slides[%d].rels != slides[%d].rels (%s != %s)' % (
      i, num, sp.rels.pprint(), cp.rels.pprint()
    )


def test_background():
  a_solidFills = prs.slides[0].background.element.xpath('//a:solidFill')
  c_solidFills = c.background.element.xpath('//a:solidFill')

  c.background.fill.solid()

  assert prs.slides[0].part.blob != c.part.blob
  assert c.background.fill.type == 1


def normpath(path):
  return os.path.normpath(os.path.join(dirname, path))

setup('test_files/test_slides.pptx')
