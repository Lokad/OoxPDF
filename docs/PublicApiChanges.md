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

The resource-admission simplification is planned separately. Its final contract
and migration instructions will be recorded here when implemented.
