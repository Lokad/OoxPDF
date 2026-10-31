These five chart inputs were saved by PowerPoint 16 after assigning its
`Chart.ChartColor` gallery values 1, 9, 10, 17 and 26 to the public
`ChartTitles/automatic-title-default.pptx` fixture. PowerPoint materializes
explicit point fills in addition to the external style parts. The tests check
those authored fills and their legend keys; they do not establish style-only
method or variation support.

The shape input places the same 24 color declarations in ordinary rectangles,
then adds eight endpoint, repeated-transform and transform-order controls.
PowerPoint's independently exported PDF supplies all expected paint values in
`office.json`; the 32 shape interiors also match its direct PNG export exactly.
Direct images remain uncalibrated against the visible application. Full-image
thresholds and any failures are recorded separately from these exact paint checks.

The JSON pins each input and reference PDF hash. Reference exports and failed
probe attempts are retained under ignored
`artifacts/chart-evaluation-1031/office-gallery/`. These are public synthetic inputs.
