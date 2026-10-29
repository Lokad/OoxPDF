# Migrating to 0.2.0

PPTX table paragraph-default bold and color now override table styling, including
an explicit `b="0"`. Direct run formatting and inherited table defaults retain
their precedence. Image deduplication uses its existing fixed identity policy;
removing its internal test hook changes no public declaration or option.

An explicit PPTX slide list now defines membership even when it yields no visible
slides. Unlisted slide relationships cannot turn an all-hidden, empty or unresolved
list into output. The existing no-page failure remains; relationship-based repair
is retained when the list is absent. No public declarations or options changed.

Unselected DOCX header/footer parts no longer produce rendering-fallback warnings
merely because a document relationship exists. Actual story references select
their diagnostic scope; selected unsupported content still warns. This corrects
false CLI strict failures without changing public declarations or picture handling.

Missing DOCX picture assets now report `IMAGE_MISSING_PART` (Error) while retaining
their authored layout space. CLI `--strict` consequently returns exit code 3 for
these incomplete packages. No public declarations or options changed; callers
should inspect diagnostics when accepting conversions with omitted picture pixels.

These intentional changes from 0.1.5 require recompiling callers. Package and
informational version are `0.2.0`; assembly/file version is `0.2.0.0`.
See [release scope and qualification](PreviewRelease.md).

Installed-font discovery also matches supported English/localized family-name
records for the same face. Canonical resolved names, program identities and
collection face indexes remain stable; alias requests retain their requested name.
DOCX Ascii/HighAnsi/EastAsia choices now apply to text fragments, with source
offsets and revision/link ownership preserved. Equivalent declared choices stay
together. These rendering repairs add no public option or shaping support.

PPTX font measurement now propagates a custom resolver's
`OoxPdfLimitExceededException` as a conversion failure, consistent with DOCX and
the fatal-limit contract. Ordinary font-I/O recovery remains available. Missing-
glyph diagnostics now exclude effectively hidden DOCX text, while visible missing
glyphs retain their warning and CLI strict status. Source metadata remains present;
these corrections add no public declarations or options.

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
The internal kerning scan-work ceiling is also removed. Repetitive tables use
cooperative cancellation; pair, coverage and class expansion limits still apply.

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

Page, PDF output and resource payload limits, plus the bounded page-content
staging window, remain. `ReportResourceUsage` still emits an informational
`CONVERSION_RESOURCE_SUMMARY`; it now reports only `pages`, `pdfPages`,
`pdfContentBytes`, `pdfOutputBytes`, `pdfFontBytes`, `pdfImageBytes` and
`pageContentSpilledBytes`. Remove dependencies on retired work/retained/peak
fields. These totals combine serialized PDF output with resource payload and spill
counts, rather than measuring managed-heap consumption. `pdfFontBytes` and
`MaxPdfFontBytesPerConversion` count uncompressed embedded-font programs and
ToUnicode data before PDF stream compression, excluding PDF dictionaries and
Standard-14 ToUnicode maps. Use `pdfOutputBytes`/`MaxOutputBytesPerConversion` for
the complete serialized PDF size. These definitions clarify existing counting;
values, limits, diagnostic IDs/fields/message and refusal behavior are unchanged.

## Optional DOCX parts

Optional styles, settings, numbering, fonts/themes and related stories now follow
the main document's owning relationships. A typed but unreferenced part does not
configure active content. Handcrafted packages that relied on the previous global
content-type fallback should add the appropriate internal relationship. Missing
or unusable optional relationships retain the corresponding empty/default behavior.
Main-document/presentation repair lookup and the public API are unchanged.

## Stream failures and shared resolvers

Seekable input may be randomly accessed. Forward-only input is supported through
bounded, cancellation-aware staging; PDF output requires only forward writes.
The direct seekable-input path avoids copying the compressed archive.

Caller-owned streams remain open on failure. Cancellation, an output limit or
an I/O failure during emission can leave an incomplete PDF prefix in the output;
discard that output before retrying. File overloads preserve an existing
destination atomically. This clarifies the existing behavior; it adds no buffering
or public API removal. Built-in resolvers now synchronize their bounded request
caches so sharing a resolver across concurrent conversions is safe.

Keep local font files unchanged while their `FileFontProgramSource` or resolver
snapshot is in use. Eviction releases a cached reference; a later request rereads
the file under the same path-based `StableId`. Previously returned bytes remain
immutable, but lifetime byte identity requires the file itself to stay unchanged.
Before replacing fonts, finish affected conversions and retire those sources and
resolvers. `WindowsFontResolver.InvalidateDiscoveryCaches()` makes subsequently
created resolvers rediscover fonts; existing resolvers keep their snapshot.
Custom `IFontProgramSource` implementations must likewise publish immutable,
byte-identical content for each identity. Bounded caching remains unchanged.

Diagnostic-observer failures now preserve the original exception identity, type
and callback stack even when its type also denotes recoverable malformed PPTX
content. They fail conversion before file publication; they are never translated
into a successful conversion with a node-failure warning. A callback that cancels
the token during the final resource summary also prevents file publication.
Callbacks remain synchronous and may invoke nested conversions. These are contract
repairs; they add no public API or buffering layer.

Asynchronous file conversions also capture the invocation working directory
before queuing. Changing the process directory while a conversion waits for a
worker no longer changes its relative input/output paths. Argument failures
still fault the returned task, and pre-cancellation still takes precedence.

Info-only diagnostic overflow now emits an informational `DIAGNOSTIC_OVERFLOW`
marker, consistent with `DiagnosticCollector.HasWarningsOrErrors == false`.
Collections that saw warnings/errors retain the error overflow marker and CLI
strict behavior, including when those entries were dropped from retained samples.
These corrections remove no public declarations.

File conversions capture absolute input/output paths before admission, resolver
or diagnostic callbacks. Changing the process working directory during a callback
does not redirect input lookup or publication. Case-only aliases of inputs,
outputs and CLI diagnostics refuse on case-insensitive mounts regardless of OS.
Provably distinct case-only names in the same directory remain allowed. Two
uncreated case-only outputs or differently spelled case-only parent paths refuse
conservatively; choose names that differ beyond case. This is a path-protection
repair with no public signature change. File-extension matching remains case-insensitive.
Directory-link ancestors are resolved before separation checks, including junctions
and symbolic-link targets with linked parents. Unresolvable/cyclic links refuse
conservatively. Final output file links remain replaceable as separate entries.
These checks do not provide protection against concurrent external directory changes.

Final input file symlinks are followed for input protection, including link chains
and targets with linked ancestors. PDF and CLI diagnostics publication reject the
readable source as well as the input entry. Separate output file links and
hardlink entries can still be replaced atomically without altering the input.
DOCX wrap searches now use cancellation instead of a 128-probe work ceiling;
long mixed-width words may fill lines that previously broke early. Unicode
boundaries, preferred breaks and the existing overwide fallback remain.
