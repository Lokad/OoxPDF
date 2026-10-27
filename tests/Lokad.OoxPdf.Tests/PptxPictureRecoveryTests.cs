using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class PptxPictureRecoveryTests
{
    public static void PptxPictureRecoveryDistinguishesMissingRelationshipAndPart()
    {
        byte[]? omitted = null;
        foreach (string variant in new[] { "present", "relationship", "part" })
            WithPicture(variant, path =>
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                using var input = File.OpenRead(path);
                using var output = new MemoryStream();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, DiagnosticSink = diagnostics.Add });
                string pdf = System.Text.Encoding.Latin1.GetString(output.ToArray());
                if (variant == "present")
                {
                    TestAssert.Equal(0, diagnostics.Count);
                    TestAssert.Contains("/Subtype /Image", pdf);
                }
                else
                {
                    TestAssert.Equal(1, diagnostics.Count);
                    OoxPdfDiagnostic diagnostic = diagnostics.Single();
                    TestAssert.Equal(variant == "relationship" ? "IMAGE_MISSING_RELATIONSHIP" : "IMAGE_MISSING_PART", diagnostic.Id);
                    TestAssert.Equal(OoxPdfSeverity.Error, diagnostic.Severity);
                    TestAssert.Equal(1, diagnostic.SlideIndex);
                    TestAssert.Equal(variant == "relationship" ? "/ppt/slides/slide1.xml" : "/ppt/media/image1.png", diagnostic.PartName);
                    TestAssert.Equal("Ignored", diagnostic.Fallback);
                    TestAssert.DoesNotContain("/Subtype /Image", pdf);
                    if (omitted is null) omitted = output.ToArray();
                    else TestAssert.True(omitted.AsSpan().SequenceEqual(output.ToArray()), "Both unavailable pictures retain the same image-free page.");
                }
            });
    }

    public static void PptxPictureRecoveryPreservesObserverFailureAndPublication()
    {
        WithPicture("relationship", path =>
        {
            string output = path + ".pdf";
            File.WriteAllText(output, "previous output");
            try
            {
                var expected = new InvalidOperationException("observer failure");
                InvalidOperationException actual = TestAssert.Throws<InvalidOperationException>(() =>
                    OoxPdfConverter.Convert(path, output, new OoxPdfOptions { DiagnosticSink = diagnostic =>
                    {
                        if (diagnostic.Id == "IMAGE_MISSING_RELATIONSHIP") throw expected;
                    }}));
                TestAssert.True(ReferenceEquals(expected, actual), "The original observer exception must propagate.");
                TestAssert.Equal("previous output", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
            }
            finally { File.Delete(output); }
        });
    }

    public static void PptxPictureRecoveryObservesCancellationAtTheDiagnostic()
    {
        WithPicture("relationship", path =>
        {
            string output = path + ".pdf";
            File.WriteAllText(output, "previous output");
            using var cancellation = new CancellationTokenSource();
            bool observed = false;
            try
            {
                TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(path, output,
                    new OoxPdfOptions { DiagnosticSink = diagnostic =>
                    {
                        if (diagnostic.Id != "IMAGE_MISSING_RELATIONSHIP") return;
                        observed = true;
                        cancellation.Cancel();
                    }}, cancellation.Token));
                TestAssert.True(observed, "Cancellation must follow the new recovery diagnostic.");
                TestAssert.Equal("previous output", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
            }
            finally { File.Delete(output); }
        });
    }

    private static void WithPicture(string variant, Action<string> check)
    {
        string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "pptx-images.pptx"));
        using var archive = ZipFile.OpenRead(source);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var stream = entry.Open(); using var bytes = new MemoryStream(); stream.CopyTo(bytes); return bytes.ToArray();
        });
        if (variant == "part") parts.Remove("ppt/media/image1.png");
        if (variant == "relationship")
        {
            const string name = "ppt/slides/_rels/slide1.xml.rels";
            XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts[name]));
            xml.Root!.Elements().Single(element => (string?)element.Attribute("Id") == "rId2").Remove();
            parts[name] = System.Text.Encoding.UTF8.GetBytes(xml.ToString());
        }
        string path = TestFixtures.WriteTempPackage(".pptx", parts);
        try { check(path); }
        finally { File.Delete(path); }
    }
}
