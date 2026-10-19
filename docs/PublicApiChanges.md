# Public API changes before the next release

Version remains 0.1.5 during implementation; release preparation is deferred.
These intentional source/binary API changes require recompiling callers.

- Removed `OoxPdfOptions.Deterministic`. It never changed conversion behavior.
  Remove the initializer; stable inputs/options still produce deterministic PDF
  bytes, and `FixedCreationDate` still controls optional document dates.
- Removed `OoxPdfOptions.Strict`. Diagnostic collection remains available through
  `DiagnosticSink`. Hosts choose their own diagnostic policy. CLI `--strict`
  continues to publish a successful conversion and return exit code 3 when
  warnings or errors were reported; ordinary success returns 0.

Removed cumulative work quotas from `OoxConversionLimits`:
`MaxChartRangeCellsPerConversion`, `MaxTableFragmentsPerConversion`,
`MaxXmlNodesPerConversion`, `MaxWorkbookCellsPerConversion`,
`MaxSceneNodesPerConversion`, `MaxNestedPackageBytesPerConversion`,
`MaxWorkbookModelsPerConversion`, `MaxImagesDecodedPerConversion` and
`MaxFontWorkPerConversion`. Use the conversion cancellation token to stop work.
Individual archive/XML/image/font bounds remain.

Removed `MaxRetainedImageBytesPerConversion`, `MaxRetainedFontBytesPerConversion`
and `MaxLiveImageBytesPerConversion`. The first two counted production without
refunds, so they did not express current memory ownership. The last estimated
image buffers alone. There is no automatic complete-memory limit.

Hosts can set `OoxPdfOptions.MemoryAdmission` to an implementation of
`IOoxPdfMemoryAdmission`. Its `TryAcquire(CancellationToken)` reserves one coarse
amount in the host's shared memory pool and returns an `IDisposable` lease, or
returns null to refuse. The host chooses the estimate for the input/workload;
include significant retained storage and construction peaks, with headroom based
on measured representative conversions. The adapter may carry that estimate;
there is no additional library quota or per-object accounting. The same pool can
coordinate concurrent conversions and other host work. Implementations must be
thread-safe; lease disposal must release exactly once and must not throw.

The library requests admission before constructing conversion storage. Refusal
throws `OoxPdfLimitExceededException` without reading the package or writing the
PDF. The lease covers parsing, font/image preparation, layout, page staging and
emission, and is disposed on success, failure and cancellation after ownership
is released. On file conversion it releases before atomic publication. Supplied
streams, resolvers and their bounded caches remain caller-owned and outside this
conversion lease. Existing bounded shared font caches also remain a shared
baseline; account for them in host pool capacity. A null option keeps ordinary conversion without host admission.

Page/output/serialized content/font/image limits and the bounded page-content
staging window remain. `ReportResourceUsage` still emits an informational
`CONVERSION_RESOURCE_SUMMARY`; it now reports only `pages`, `pdfPages`,
`pdfContentBytes`, `pdfOutputBytes`, `pdfFontBytes`, `pdfImageBytes` and
`pageContentSpilledBytes`. Remove dependencies on retired work/retained/peak
fields. These totals describe output and spilling, not managed-heap consumption.
