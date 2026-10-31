`office.json` retains independently measured PowerPoint 16 plot edges and source/
reference hashes. Controls vary title absent/automatic/explicit, legend position,
series count, stacked grouping, frame width, and axis font size/label side.
Original inputs, Office PDFs and parent outputs remain in ignored
`artifacts/tplot-frame-1031/`.

Office keeps horizontal plot edges unchanged when titles are added. Titles and
top/bottom legends consume vertical space. The portable tests challenge this
invariant across pinned fonts, frame widths, series counts and legend types;
hidden and opposite-side labels also retain the measured edge-padding contract.
The raw Office coordinates remain separate from font-dependent text measurements.

Manual and side-legend controls retain their explicit placement fallbacks. Their
existing inaccuracies are not corrected by these tests. Whole-image thresholds
and remaining typography/vertical-layout/legend failures remain separate from
the horizontal-reserve checks. Direct application-view calibration is pending.
