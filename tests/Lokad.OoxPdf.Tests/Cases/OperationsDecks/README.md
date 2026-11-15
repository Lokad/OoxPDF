# Independent operations decks

`tools/NewOperationsDeckFixtures.py` writes these public synthetic presentations
from handwritten OOXML, including their own minimal theme, master and layout.
They reuse no slides or chart definitions from the original portfolio.

`service-operations.pptx` uses an 800×600pt 4:3 frame and right legend. Held-out
`survey-operations.pptx` uses a 1000×500pt 2:1 frame, bottom legend and adjacent
prose. Both have two slides, an explicitly colored six-series chart with two
categories, a generated opaque status picture on each slide, ordinary text and
an external hyperlink on the checklist slide. Values are entirely illustrative;
the cache-only chart declares all values directly in its XML.

`expected.json` pins source hashes, authored slide dimensions, every intended
legend/category/title/prose string, picture rectangles and link counts. All six
names must paint once on the chart slide; a whole-page score cannot substitute
for their presence. The authored chart, prose and pictures must remain within
their own frames. These are content/composition contracts, not independent Office
color, chart-placement or application-view calibration claims.
