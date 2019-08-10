import os
from pptx import Presentation as P

dirname = os.path.dirname(__file__)
try:
  import pptxy
except ImportError:
  import sys
  sys.path.insert(0, os.path.normpath(os.path.join(dirname, '..')))
  import pptxpy


p = P(os.path.normpath(os.path.join(dirname, 'test_files/test_slides.pptx')))
slides_num = len(p.slides)
c = p.slides.duplicate(0)

assert len(p.slides) == slides_num + 1
assert p.slides[-1] is c
assert c.part.partname == '/ppt/slides/slide2.xml'
assert p.slides[0].part.blob == c.part.blob

assert p.slides[0].part.rels == c.part.rels

a_solidFills = p.slides[0].background.element.xpath('//a:solidFill')
c_solidFills = c.background.element.xpath('//a:solidFill')

c.background.fill.solid()

assert p.slides[0].part.blob != c.part.blob
#assert p.slides[0].background.fill.type == 5
assert c.background.fill.type == 1
#assert len(a_solidFills) != len(c_solidFills)
