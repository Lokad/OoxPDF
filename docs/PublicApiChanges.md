# Caller migration

## 0.3.0

`WindowsFontResolver.CreateAsync(cancellationToken)` prepares installed-font
metadata with awaited file reads. `Resolve` and catalog queries then use the
published snapshot without file I/O; program bytes remain lazy. The existing
constructor now completes metadata discovery during construction instead of the
first query, blocking on the same initialization core. Concurrent callers share
one snapshot; canceled initialization publishes nothing, and canceled waiters do
not cancel other callers. Cache invalidation still affects only future resolvers.
Directory enumeration has no asynchronous platform equivalent here.

`IFontProgramSource.GetBytesAsync` is replaced by
`AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)`,
returning an `IOoxPdfFontProgramLease`. Custom sources must implement the new
contract; keep immutable `lease.Bytes` valid until disposal and finish every I/O
borrow before refunding storage. The converter releases each source lease after
the parser takes its owned copy. Caller-funded memory sources need no second copy
or reservation. See [font ownership](Fonts.md).

`OoxPdfFontPackResolver.CreateHttpAsync` accepts an optional `cacheBudget` before
the cancellation token. Migrate positional fourth-token calls to the named
`cancellationToken:` argument. The resolver and owned `FileFontProgramSource`
instances are disposable. Supply a longer-lived cache budget for governed shared
retention; otherwise a governed operation funds private acquisition. Cache
shutdown/eviction ends retention without invalidating active leases. Keep its
borrowed budget alive through pending acquisitions and the last lease. The
resolver never disposes the caller's `HttpClient` or budget.

`ConvertAsync` awaits bounded package input, ZIP extraction, page staging and PDF
output, including owned storage disposal. Synchronous entry points block on the
same pipeline. DOCX awaits font-program acquisition during preparation; measurement
and painting then use prepared fonts. Default local discovery awaits font-file
metadata reads before either renderer starts; directory enumeration remains a
synchronous platform operation. PPTX awaits declared shape/table paragraph
fonts, generated bullets/chart text, ellipsis and needed glyph fallback candidates
before measurement, using the shared prepared font owner. Caller streams
stay open; cancellation waits for outstanding
I/O before refunding borrowed storage. Input failures escape without preflight retry.

Set `TemporaryStorage` when embedded stream conversions may exceed
`MaxResidentPageContentBytesPerConversion`. Without it, exceeding that existing
window fails explicitly and creates no process-temp files. File/CLI convenience
conversions select filesystem backing. To choose a directory, use
`new OoxPdfFileTemporaryStorage(directory)`. Custom `IOoxPdfTemporaryStorage`
providers return fresh empty readable/writable/seekable streams. The library owns
and asynchronously disposes each stream. Providers borrow the supplied nullable
budget, fund significant private storage/buffers and refund after their last user;
they must not dispose or retain that budget after storage closes.

`PDF_PAGE_CONTENT_SPILLED` keeps its ID and informational severity; its fallback
is now `temporary-storage`, covering host-selected backing as well as files.

PPTX table minimum-row sizing now uses the supplied font resolver and cancellation
token, matching cell layout and painting. Explicit custom fonts can therefore
change minimum row heights that previously depended on ambient system fonts.

An explicit `FontResolver` now defines the entire allowed font universe for PPTX
as well as DOCX. Custom resolvers without a built-in catalog no longer gain
implicit Windows fallback fonts. Provide all required glyph coverage through
the resolver or a font pack; missing glyphs retain the diagnosed visible
replacement. Null still selects local Windows discovery. This can change PPTX
output that previously depended on ambient fonts despite supplying a custom resolver.

`OoxPdfLimitExceededException` now exposes `Kind` and `Scope`. Host admission
refusal reports `MemoryAdmission` / `Conversion`; the existing constructors retain
`DocumentBound` / `Conversion`. Custom shared-resource providers can use the new
constructor with `SharedResource`. Existing exception catches still work. Refusal
remains terminal and preserves file publication and caller-stream ownership.
Admission coverage remains partial; see the [operating envelope](OperatingEnvelope.md).
Conversion now awaits input, temporary storage, output and selected DOCX/PPTX font
programs. CPU layout and drawing run synchronously between those awaits.

Replace `OoxPdfOptions.MemoryAdmission` / `IOoxPdfMemoryAdmission` with
`MemoryBudget` / `IMemoryBudget`. The caller owns the shared root; the library
creates zero-reservation scopes and asks producers to reserve and resize their
capacity before allocation, including overlapping buffers during growth. Scopes
own their own reservations, share the root allowance, and do not dispose the root
or other independently created scopes. Dispose/refund only after storage users
finish. Refusal never triggers eviction, spilling or reduced output.

Producer admission covers package intake and coarse parsed XML/cache capacity, extracted parts and copy buffers,
including nested chart workbooks, plus staged PDF content and its fixed encoding/
readback buffers. Providers own their private backing/buffer admission; the built-in
file provider adds no duplicate managed file buffer. Package bytes and resident staged content remain funded
through emission; temporary intake buffers and transferred resident bytes refund
after their users finish. Private copied/extracted SFNTs and subset program buffers
are also admitted before construction, held through emission, and refunded after
failed preparation. Only the selected collection face is copied and charged.
PNG/BMP decoded planes are admitted with PNG IDAT/scanline construction overlap;
temporary storage refunds after decode and planes after transformation/compression
users finish. PNG palette/transparency slices borrow their source instead of
copying it. Baseline JPEG decoding also funds padded component/RGB-plane overlap;
only RGB storage escapes with its pixel owner. Ordinary JPEG passthrough borrows
encoded bytes. Image/font compression admits growing-output capacity, replacement
overlap and final copying without source-sized preallocation. Image payloads stay
funded through emission; font-compression buffers refund after their stream is
written. Crop and generated shadow/glow planes are admitted for their invocation;
recolor reuses private pixels and leaves encoded input/cache bytes unchanged.
Other parsed maps, local-discovery metadata and PDF resources still need producer migration.
This partial coverage is not a complete conversion-memory limit. Null budget
means ungoverned memory; input/output bounds and cancellation remain independent.
Derived document/scene/workbook models now request their own coarse capacity before
construction and retain it through their producer's last reader. Embedded-workbook
models survive their temporary package's disposal. That model group refunds after
page production, before serialization; source buffers and PDF resources retain
their separate emission lifetimes. Existing budgeted inputs can refuse earlier.
Repeated DOCX page/static-story placement now requests reusable coarse capacity
before construction, including overlapping candidate/frame copies. Repeated body-
table headers include nested source content and column multiplicity in that same
grant; it grows before a continuation header is copied. It shares the
model producer lifetime and does not charge cumulatively per retry. Existing host
allowances may therefore refuse documents previously accepted without this storage
being admitted. A font pack's existing `cacheBudget` now also admits coarse
manifest-derived metadata before parsing/indexing and holds it until shutdown
clears the indexes and pending acquisitions finish. Live program leases retain
separate byte funding. Returned catalog values are caller-owned; their sources
reject new acquisition after shutdown. Small cache allowances can therefore refuse
pack creation earlier, with `MemoryAdmission` / `SharedResource` rather than a
transport error or retry. Local-discovery metadata, other parsed maps and PDF/vector
resource coverage remain partial. SVG now requests coarse render-local DOM/path/
gradient working capacity before parsing, then refunds it after painting/failure.
A host refusal escapes DOCX raster fallback and PPTX ignored-node recovery; small
budgets can refuse earlier. This does not fund retained PDF/vector resources or
change the rendering rules. The removed estimated lease is not a second memory model.

The package provides .NET 10 and .NET 11 assets. Common Strict DOCX content now
renders through the existing pipeline. `OOXML_STRICT_DIALECT` remains a warning,
with DOCX `Fallback` changed to `Partial`; callers must still inspect other
unsupported-content diagnostics. Strict PPTX remains unsupported. No public
declarations or conversion options were added for this support.

## Migrating to 0.2.0

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

In **0.2.0**, hosts could set `OoxPdfOptions.MemoryAdmission` to an implementation of
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

Asynchronous file conversions capture the invocation working directory before
awaiting input or invoking host callbacks. Later directory changes cannot redirect
relative input/output paths. Argument failures still fault the returned task,
and pre-cancellation still takes precedence.

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

Staged vector resources now share the content store's memory reservation through
emission. A governed conversion may refuse earlier when its descriptors, shading
stops or retained patterns cannot fit; spilling content does not refund still-live
resources. No public signature changes. Current-page painting/construction remains
partial; this is coarse admission, not complete heap accounting.

DOCX/PPTX painting now requests one coarse construction reservation per page
before graphics creation. A host may refuse before page content is built; the
reservation survives pending staging I/O and refunds when page production advances
or is disposed. The same host capacity can serve later pages. Staged descriptors
retain separate funding through emission. No new option or public signature.
