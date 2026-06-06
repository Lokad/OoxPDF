# Lokad.OoxPdf.Tests

The test project is a dependency-free console runner. Tests are cataloged by capability group so the PPTX
renderer can be tightened bottom-up without running the whole suite on every iteration.

```powershell
dotnet run --project tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj -- --list
dotnet run --project tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj -- --skip-slow
dotnet run --project tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj -- --group pptx-typography --skip-slow
```

Primary groups are `pptx-model`, `pptx-typography`, `pptx-shapes`, `pptx-images`,
`pptx-composition`, `pptx-tables`, `pptx-charts`, and the corresponding `docx-*` groups.
