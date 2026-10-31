`office.json` records measured geometry from 24 independently opened PowerPoint 16
controls: horizontal/vertical bars, one through four series, and overlap -27/0/+50
with a 219% gap. Each reference contains six bars per series. Inputs derive from
the public `AuthoredColorTransforms/gallery-17.pptx` fixture by cloning cached
series, assigning distinct indices, and changing the overlap and axis direction.
Input and independent PDF hashes are retained with the raw measurements.

The test compares painted bar thickness and interseries steps relative to the
category band, at the existing 0.01pt geometry tolerance. It also retains the two
historical three-series measurements from the removed private-helper test. This
qualifies cluster spacing; automatic plot-frame, axis and legend placement remain
approximate. Whole-image gates and their unchanged thresholds are separate.

Original inputs, PDFs and native/candidate measurements remain under ignored
`artifacts/tbar-spacing-1031/v2/`. No new Office fidelity claim follows from the
historical measurements alone.
