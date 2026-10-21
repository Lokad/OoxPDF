# PDFium for Windows validation

The separate managed PDFium rasterizer loads the vendored native
bin/pdfium.dll. Its SHA-256 is pinned in pdfium.sha256; RasterizePdf.ps1
checks the pin before running. VERSION records the upstream binary version;
LICENSE and licenses/ retain its license notices.

Run from the repository root:

    pwsh tools/RasterizePdf.ps1 -InputPdf input.pdf -OutputDirectory artifacts/raster

This builds only the validation tool when needed. Use -SkipBuild to reuse its
existing Debug build. No pdfium_test.exe, native headers or import library are
required. PDFium remains outside the shipped Lokad.OoxPdf library.
