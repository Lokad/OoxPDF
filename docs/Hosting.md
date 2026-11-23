# Hosting and resource ownership

The library is safe managed code with no native runtime dependencies. It uses
cooperative cancellation, format bounds, bounded caches and page staging. These
controls do not guarantee conversion of arbitrary documents or cap the whole heap.

## Output and staging

File conversion stages output, cleans owned storage deterministically and
publishes atomically. Existing destinations survive a failure. Supplied streams,
resolvers, budgets and temporary-storage providers remain caller-owned. Stream
failures can leave a prefix; discard it.

Cleanup is awaited even after failure. The original exception remains primary;
a secondary disposal exception is available through
`exception.Data["Lokad.OoxPdf.CleanupException"]` (an `AggregateException` if several
cleanup faults accumulated). A disposal failure alone still fails conversion.

Page/group content has a resident staging window. Stream conversion exceeding it
requires `OoxPdfOptions.TemporaryStorage`; the library does not silently use the
process temp folder. File/CLI conversion selects filesystem backing. A custom
`IOoxPdfTemporaryStorage.CreateAsync` returns a fresh empty readable/writable/
seekable stream positioned at zero. The converter owns and asynchronously disposes
that stream. The provider supports concurrent calls and borrows the supplied budget.
`OoxPdfFileTemporaryStorage` is the built-in filesystem provider.

## Shared memory admission

Set `OoxPdfOptions.MemoryBudget` to a host implementation of `IMemoryBudget` for
coarse shared-memory admission. It covers significant package/model construction,
selected font/image/vector resources, page construction and staged content.
Coverage remains partial; leave headroom for incidental objects and collection lag.
No option represents complete heap accounting.

- `CreateScope()` creates an independent reservation group against the same pool;
  it reserves nothing by itself and does not dispose the caller's root budget.
- `TryReserve(bytes, out reservation)` admits capacity before acquisition, without
  waiting. Refusal returns false/null without changing the pool.
- `IMemoryReservation.TryResize(bytes)` retains the old amount on refusal;
  shrinking succeeds. Growth must cover old/new buffer overlap until old readers finish.
- Scope/reservation disposal refunds exactly once after final use, including
  cancellation/failure, and must not throw. Providers must be thread-safe.

Refusal throws `OoxPdfLimitExceededException` with `MemoryAdmission` and the
appropriate conversion/shared-resource scope. It terminates conversion rather than
triggering reduced output, an unrelated fallback or a transport retry. Null budgets
leave memory ungoverned while format/output bounds and cancellation still apply.
Long-lived font caches use separately supplied cache budgets; see [fonts](Fonts.md).
Caller-owned input/output buffers need caller-owned funding through their final use.

## Output bounds

`OoxConversionLimits` exposes these nonnegative limits:

| Property | Default |
|---|---|
| `MaxPagesPerConversion` | 10,000 |
| `MaxPdfContentBytesPerConversion` | 1 GiB |
| `MaxOutputBytesPerConversion` | 2 GiB |
| `MaxPdfFontBytesPerConversion` | 256 MiB |
| `MaxPdfImageBytesPerConversion` | 512 MiB |
| `MaxResidentPageContentBytesPerConversion` | 64 MiB |

Font totals count uncompressed embedded programs, ToUnicode and CID-to-glyph maps before
compression, excluding dictionaries and standard Type1 maps. The complete PDF
size is the output-byte total. Archive/XML/image/font bounds apply independently.
These are payload limits, not cumulative work quotas; use cancellation for work.

`ReportResourceUsage` emits `CONVERSION_RESOURCE_SUMMARY`: `pages`, `pdfPages`,
`pdfContentBytes`, `pdfOutputBytes`, `pdfFontBytes`, `pdfImageBytes` and
`pageContentSpilledBytes`. Stream summaries report page, content and spill totals
before serialization, with zero output, font and image bytes at that point.
File summaries include completed serialization and precede atomic publication.
These counters measure PDF payloads and staging. A throwing observer aborts conversion.

Input, font, temporary-storage and output I/O are awaited; layout and drawing are
synchronous. On .NET 10, bounded async intake copies even seekable input before
the BCL's synchronous ZIP directory materialization. .NET 11 can read seekable
input directly. Choose concurrency and admission from representative host workloads;
no universal throughput or maximum-document-size claim is made.
