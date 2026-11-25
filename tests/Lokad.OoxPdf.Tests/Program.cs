using Lokad.OoxPdf.Tests;

return TestTemporaryDirectory.Run(() => TestRunner.Run(args, TestCatalog.All, typeof(TestCatalog).Assembly));
