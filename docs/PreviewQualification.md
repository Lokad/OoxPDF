# 0.2.0-preview qualification — 2026-10-20

The third bounded review is locally qualified and the preview package is prepared.
Publication, tagging, release creation, renderer optimization and broader feature
expansion remain deferred. See [scope](PreviewRelease.md) and [migration](PublicApiChanges.md).

One Release runtime is frozen from `04667ad36298dddfb0c3f5b49f1d4f243a157174`.
Library SHA-256:
`394AB3044AAA80060F88B6648CFFB21B3FE1B65EE2BDA926867B9174D779814F`.
Test assembly SHA-256:
`5ECA094764619A5E1D4202ED0C3E726BBBB90AC54124FEB351EC6647E2BDA898`.
[Candidate CI](https://github.com/Lokad/OoxPDF/actions/runs/39468596965) uses
`3ab20f2a`, which adds the inspector to the solution without changing library/tests.
Its independently built Windows/Ubuntu outputs are separate from this artifact.
Both Windows and Ubuntu jobs passed. Final documentation changes leave the
qualified source and assets unchanged; no broad suite is replayed.

| Gate | Result |
|---|---|
| Affected checks | All 41 pass, extracted from the Windows milestone report; no affected-suite replay |
| Windows milestone | 3,732 pass, **1 inherited fail**, 3 environmental skips; includes the two retained exhaustive balloon methods and four CLI compatibility checks |
| Same frozen DLL on Ubuntu | 822 pass, 0 fail, 49 environmental skips; one exact 871-method invocation, no library rebuild |
| Parent CLI controls | Both new alias regressions fail against the preceding CLI; separate failure diagnostics pass |
| Qualification tools, each platform | 23 freeze checks, 10 build guards, one report join/seven refusals, 10 frozen bindings and 8 inspector-build checks |
| Capability catalog / manifests | All 275 capabilities and 2,183 distinct selectors resolve; 715 visual manifests validate. These checks execute no rendering tests and do not imply visual fidelity |
| Frozen markup integration | Actual DOCX inspection on both platforms; complete cached-markup smoke with matching converter/inspector and no build. Its reference is the preceding candidate, an identity control without an Office parity claim |
| Package consumers | Windows and Ubuntu consume the same final .nupkg; DLL identity, version, caller guides, embedded fonts and admission lifetime pass |
| Representative documents | Nine sources, sixteen outcomes, 32 pages; every PDF, diagnostic list and recorded semantic result matches the accepted parent |
| Independent Office comparisons | Reused by input/options/PDF identity: two PPTX strict passes, fourteen DOCX strict failures; no export, new raster or threshold change |
| Root transfer | All 24 assets verified and two integration PDFs match; no root build or suite replay |
| Safe managed | Explicit AllowUnsafeBlocks=false, IL-only, zero P/Invoke/imported COM types, no NuGet runtime dependencies |

The local balloon assertion sees a 12.61pt single-row body where its Office-derived
expectation requires a second row. The preceding qualified DLL fails identically
on this machine; its cause remains unresolved. The assertion and rendering guards
stay unchanged. Previous CI skipped it without Aptos. Local qualification retains
this failure even when affected checks pass.

## Bounded changes

CLI diagnostics-path aliases are rejected before input detection can fail, preserving
both source and existing destination for unsupported, malformed or missing inputs.
Ordinary separate failure diagnostics and strict-exit behavior remain covered.
Transfer and report reuse now include root build settings in existing revision checks;
dirty/committed props and newly added targets cannot silently satisfy transfer identity.

The two exhaustive balloon assertion sweeps remain available as slow/milestone tests.
They pass once in 159.813 seconds combined. Three focused methods cover layout boundaries and
the live spacing predicate in 422 ms on the same Windows run. Rendering policy and
Office assertions are unchanged; this reduces routine qualification cost, not renderer
runtime. Capability selections retain neighboring checks and explicit exhaustive access.

Cached markup qualification accepts a frozen CLI and matching DOCX inspector together,
checks library hashes and forwards no-build binding. DOCX/PPTX/PDF inspectors accept
explicit runtimes; independent validation tools remain separate. The DOCX inspector
is included in the one frozen artifact and solution build. Deleted exactly two tiny
test-only helpers, IsEmptyMatrix and IsSupportedSimpleCell; tests use live measurement
or local source-shape checks. Capability prose now describes qualified composition
and explicit fallbacks. No public API removal or renderer behavior change is introduced.
Earlier cleanup and intentional API removals remain in the migration guide and Git
history; the preceding qualification is retained under artifacts/review-second-20261020/.

## Packages and residuals

Final packages: `artifacts/nuget/preview-020-third-review-20261020/`.

- .nupkg SHA-256: `2B09FE064A3634AEC4F62874BFE25D599FB888063ED563C9C4F6B77D8C4BD241`.
- .snupkg SHA-256: `0237199FA3AD8455C8F5CAF61ECC02046E9A7E6D137960D0D051B6E75ADFFF02`.

Packed from the committed frozen build without recompilation. All four bundled caller
guides match its checkout and normalized Git text. Package/informational version stays
`0.2.0-preview`; assembly/file version stays `0.2.0.0`.

Fourteen DOCX strict failures include significant geometry and typography. SVG Review
prose Y remains 1.436035pt against the unchanged 1pt gate; raw linked-footnote Word URI
difference remains 46.794pt while native picture/link bounds agree within .000016pt.
Two cached-field/extraction flags remain explicit. The fixed portfolio shows no new
supported-scope content or pagination regression; it does not establish arbitrary-
document ink/clipping fidelity. Chart point-count, exact-five-paragraph balloon and
font/document-shape guards retain unresolved behavior without expansion or replacement.

Full identities, deletions, controls, failed attempts and residuals:
`artifacts/review-third-20261020/qualification.json`. Latest scan: **5.044 GB**, zero errors; near the loose cap. Historical evidence
and runtimes remain preserved.
Preserve safe managed execution, input bounds, coarse admission/refunds, bounded
caches/page staging, cancellation, deterministic cleanup, atomic publication and PDF
validity. Keep the loose 5 GB workspace cap and infrequent checks.
