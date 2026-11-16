# Migrating to the current API

The current 0.3.0 source API requires recompilation of callers built against 0.2.0,
including separately built components. Published 0.2.0 remains a separate package.
This page lists compatibility actions; current ownership contracts are in
[hosting](Hosting.md), [fonts](Fonts.md) and [usage](Usage.md).

| Previous use | Current API |
|---|---|
| `MemoryAdmission` / `IOoxPdfMemoryAdmission` | Supply `MemoryBudget` / `IMemoryBudget`; keep the caller-owned root alive through all scopes/reservations. |
| `IFontProgramSource.GetBytesAsync` | Implement/use `AcquireAsync(operationBudget, ct)` and retain the immutable lease through its final reader. |
| Precompiled `new FileFontProgramSource(path)` | Recompile; the optional cache-budget parameter changes the binary constructor signature. |
| Fourth positional token in `CreateHttpAsync` | Use `cancellationToken:` because the fourth parameter is the optional cache budget. |
| Owned file/font-pack/owned Windows resolver | Dispose after use; keep borrowed cache funding alive until outstanding leases finish. |
| Stream conversions needing staging beyond the resident window | Supply `TemporaryStorage`; file conversions select filesystem backing. |
| Custom PPTX resolver relying on ambient system fonts | Supply the complete required font universe through the resolver or font pack. |

Host admission can refuse covered storage earlier than an ungoverned conversion;
coverage remains coarse and partial. Limit exceptions expose `Kind` and `Scope`.
Cancellation and refusal remain terminal and preserve file publication contracts.

Do not infer comment presence from an unsupported-comments warning: Final/Original
and document-hidden comments are intentionally unprinted. Select the required
markup view, and inspect diagnostics for visible unsupported content.

For callers predating 0.2.0, remove `OoxPdfOptions.Deterministic` and library `Strict`;
use stable inputs/fonts and host diagnostic policy. CLI `--strict` remains available.
Removed cumulative-work and nonrefunding retained-byte limit properties have no
replacement fuel system. Use cancellation, the current output bounds and optional
shared admission. Recompile rather than treating these changes as binary-compatible.
