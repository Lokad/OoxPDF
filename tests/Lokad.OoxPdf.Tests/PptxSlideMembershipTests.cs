using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSlideMembershipTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void PptxSlideMembershipMatchesIndependentOfficeControls()
    {
        foreach (var (id, count) in new[] { ("ordered-visible", 2), ("unlisted-visible", 1),
            ("listed-hidden-unlisted-visible", 0), ("listed-hidden-no-orphan", 0) })
        {
            PptxDocument document = Read(Case(id));
            TestAssert.Equal(count, document.Slides.Count);
            for (int i = 0; i < count; i++)
            {
                TestAssert.Equal("/ppt/slides/slide" + (i + 1) + ".xml", document.Slides[i].PartName);
                TestAssert.Equal(i, document.Slides[i].Index);
            }
        }
    }

    public static void PptxSlideMembershipExplicitEmptyAndUnresolvedListsStayEmpty()
    {
        foreach (bool unresolved in new[] { false, true })
            WithPresentation("ordered-visible", xml => {
                XElement list = xml.Root!.Element(P + "sldIdLst")!;
                list.RemoveNodes();
                if (unresolved) list.Add(new XElement(P + "sldId", new XAttribute("id", "256"),
                    new XAttribute(R + "id", "missingRelationship")));
            }, path => TestAssert.Equal(0, Read(path).Slides.Count));
    }

    public static void PptxSlideMembershipRetainsSourceOrderAndMissingListRecovery()
    {
        WithPresentation("ordered-visible", xml => {
            XElement list = xml.Root!.Element(P + "sldIdLst")!;
            list.ReplaceNodes(list.Elements().Reverse().ToArray());
        }, path => {
            PptxDocument document = Read(path);
            TestAssert.Equal("/ppt/slides/slide2.xml", document.Slides[0].PartName);
            TestAssert.Equal("/ppt/slides/slide1.xml", document.Slides[1].PartName);
            TestAssert.Equal(0, document.Slides[0].Index);
            TestAssert.Equal(1, document.Slides[1].Index);
        });
        foreach (bool hidden in new[] { false, true })
            WithPresentation(hidden ? "listed-hidden-unlisted-visible" : "ordered-visible",
                xml => xml.Root!.Element(P + "sldIdLst")!.Remove(), path => {
                    PptxDocument document = Read(path);
                    TestAssert.Equal(hidden ? 1 : 2, document.Slides.Count);
                    TestAssert.Equal(hidden ? "/ppt/slides/slide2.xml" : "/ppt/slides/slide1.xml", document.Slides[0].PartName);
                    TestAssert.Equal(0, document.Slides[0].Index);
                });
    }

    public static void PptxSlideMembershipEmptyOutputPreservesPublicationAndCancellation()
    {
        string directory = Path.Combine(Path.GetTempPath(), "slide-membership-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string output = Path.Combine(directory, "output.pdf");
            foreach (string id in new[] { "listed-hidden-unlisted-visible", "listed-hidden-no-orphan" })
            {
                File.WriteAllText(output, "retained previous output");
                ArgumentException failure = TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(Case(id), output));
                TestAssert.Contains("at least one page", failure.Message);
                TestAssert.Equal("retained previous output", File.ReadAllText(output));
                File.Delete(output);
                TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(Case(id), output));
                TestAssert.Equal(false, File.Exists(output));
                using var cancelled = new CancellationTokenSource();
                cancelled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => Read(Case(id), cancelled.Token));
                TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(Case(id), output, cancelled.Token));
                TestAssert.Equal(0, Directory.GetFiles(directory).Length);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static string Case(string id) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "SlideMembership", id + ".pptx"));

    private static PptxDocument Read(string path, CancellationToken token = default)
    {
        using var input = File.OpenRead(path);
        return new PptxReader().Read(OoxPackage.Open(input, default), token);
    }

    private static void WithPresentation(string id, Action<XDocument> mutate, Action<string> check)
    {
        string path = Path.Combine(Path.GetTempPath(), "slide-membership-" + Guid.NewGuid().ToString("N") + ".pptx");
        File.Copy(Case(id), path);
        File.SetAttributes(path, FileAttributes.Normal);
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("ppt/presentation.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                mutate(xml);
                entry.Delete();
                using var output = archive.CreateEntry("ppt/presentation.xml").Open();
                xml.Save(output);
            }
            check(path);
        }
        finally { File.Delete(path); }
    }
}
