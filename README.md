[![Travis](https://travis-ci.org/denim2x/pptx-py.svg?branch=master)](https://travis-ci.org/denim2x/pptx-py)

*pptx-py* is a Python library with various tools for enhancing [python-pptx](http://github.com/scanny/python-pptx).

## Features
- `Slides.duplicate(index|id)`
  given the `Slide` instance by `index` or `id` from the current `Slides` instance,
  creates an *exact* copy of its contents by cloning its underlying XML structure,
  wrapping it in a new `SlidePart` instance and inserting it in the corresponding 
  `PresentationPart` instance and, thus, in the current `Slides` instance

## Dependencies
Currently there are no *explicitly* required external dependencies, though in order
to use this library, the presence of **python-pptx** needs to be ensured.

## License
MIT License
