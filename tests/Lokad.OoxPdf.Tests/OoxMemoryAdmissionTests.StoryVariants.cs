using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void LayoutRelatedStoryVariantsReuseCapacityAndSurvivePackageDisposal()
    {
        var measurer = LayoutAdmissionMeasurer();
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        using (var input = RelatedStoryVariantInput())
        {
            var package = OoxPackage.Open(input, default);
            var source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            var measurements = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
            int additionalFamilies = 0;
            pool.RefuseChange = (before, after) => {
                if (before == 0 && after > 0 && measurements.Count > 0) additionalFamilies++;
                return false;
            };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                { PlacementRelatedStoryMeasurements = measurements };
            var first = engine.Create(source, measurer, default);
            TestAssert.Equal(8, measurements.Count);
            TestAssert.Equal(7, additionalFamilies);
            TestAssert.Equal(8, first.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.IsNormalStoryType).Select(note => note.StoryLayout.Story.Id).Distinct().Count());
            var cached = measurements.ToArray();
            long held = pool.Reserved;
            pool.Capacity = held;
            var second = engine.Create(source, measurer, default);
            TestAssert.Equal(held, pool.Reserved);
            TestAssert.Equal(7, additionalFamilies);
            TestAssert.Equal(first.Pages.Count, second.Pages.Count);
            TestAssert.True(cached.All(pair => ReferenceEquals(pair.Value, measurements[pair.Key])),
                "Placement retries borrow the same admitted width families.");
            package.Dispose();
            TestAssert.True(pool.Reserved > 0, "Measured variants outlive package storage.");
            TestAssert.Contains("Variant note", measurements.Values.Last().Last().TextLines[0].Text);
            TestAssert.True(second.Pages.SelectMany(page => page.PlacedRelatedStories).Any(), "The reused layout remains readable.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void LayoutRelatedStoryVariantFailureRefundsUnpublishedFamily()
    {
        // Refusal, cancellation at admission, and cancellation during construction.
        foreach (int failure in new[] { 0, 1, 2 })
        {
            var inner = LayoutAdmissionMeasurer();
            var pool = new TestMemoryBudget();
            using var cancellation = new CancellationTokenSource();
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            using (var input = RelatedStoryVariantInput())
            using (var package = OoxPackage.Open(input, default))
            {
                var source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
                var measurements = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
                bool reached = false;
                long prior = 0;
                pool.RefuseChange = (before, after) => {
                    if (reached || before != 0 || after <= 0 || measurements.Count != 1) return false;
                    reached = true;
                    prior = pool.Reserved;
                    if (failure == 1) cancellation.Cancel();
                    return failure == 0;
                };
                var measurer = new RelatedStoryAdmissionMeasurer(inner, () => {
                    if (failure == 2 && reached) { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
                });
                var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    { PlacementRelatedStoryMeasurements = measurements };
                if (failure == 0)
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        engine.Create(source, measurer, cancellation.Token)).Kind);
                else
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
                        engine.Create(source, measurer, cancellation.Token)).CancellationToken);
                TestAssert.True(reached, "A second width requests capacity before measuring its family.");
                TestAssert.Equal(1, measurements.Count);
                TestAssert.Equal(prior, pool.Reserved);
                var retained = measurements.Single();
                pool.RefuseChange = null;
                var retry = engine.Create(source, inner, default);
                TestAssert.Equal(8, measurements.Count);
                TestAssert.True(ReferenceEquals(retained.Value, measurements[retained.Key]), "Failure leaves the earlier family reusable.");
                TestAssert.Equal(8, retry.Pages.SelectMany(page => page.PlacedRelatedStories)
                    .Where(note => note.StoryLayout.Story.IsNormalStoryType).Select(note => note.StoryLayout.Story.Id).Distinct().Count());
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void LayoutRelatedStoryVariantPublicRefusalPreservesPublication()
    {
        long request = 0;
        var observer = new TestMemoryBudget();
        var measurements = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
        var measurer = LayoutAdmissionMeasurer();
        using (var operation = OoxConversionBudget.BeginScope(null, observer))
        using (var input = RelatedStoryVariantInput())
        using (var package = OoxPackage.Open(input, default))
        {
            var source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            observer.RefuseChange = (before, after) => {
                if (before == 0 && after > 0 && measurements.Count == 1) request = after;
                return false;
            };
            _ = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                { PlacementRelatedStoryMeasurements = measurements }.Create(source, measurer, default);
        }
        TestAssert.True(request > 0, "Observe the live extra-width request, independently of other producers.");
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-story-refusal-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            using var bytes = RelatedStoryVariantInput();
            string source = Path.Combine(directory, "source.docx"), target = Path.Combine(directory, "result.pdf");
            File.WriteAllBytes(source, bytes.ToArray());
            foreach (bool file in new[] { false, true })
            {
                bool reached = false;
                var pool = new TestMemoryBudget { RefuseChange = (before, after) => {
                    if (before != 0 || after != request) return false;
                    reached = true;
                    return true;
                }};
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryBudget = pool,
                    FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })) };
                File.WriteAllBytes(target, [1, 2, 3]);
                using var input = new MemoryStream(bytes.ToArray());
                using var output = new MemoryStream();
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() => {
                    if (file) OoxPdfConverter.Convert(source, target, options);
                    else OoxPdfConverter.Convert(input, output, options);
                }).Kind);
                TestAssert.True(reached, "Public conversion reaches the extra-width request.");
                TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 1, 2, 3 }), "Refusal preserves the destination.");
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(input.CanRead && output.CanWrite, "Caller streams remain open.");
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private sealed class RelatedStoryAdmissionMeasurer(IDocxTextMeasurer inner, Action beforeMeasure) : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            beforeMeasure();
            return inner.MeasureText(run, text, fontSize);
        }
    }

    private static MemoryStream RelatedStoryVariantInput()
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string Section(int index) => $"<w:sectPr><w:pgSz w:w=\"{12240 + index * 120}\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr>";
        string body = string.Concat(Enumerable.Range(0, 8).Select(index =>
            $"<w:p><w:pPr>{(index < 7 ? Section(index) : "")}</w:pPr><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id=\"{index + 1}\"/></w:r></w:p>"));
        return TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\"><w:body>{body}{Section(7)}</w:body></w:document>",
            ["word/endnotes.xml"] = $"<w:endnotes xmlns:w=\"{w}\">" + string.Concat(Enumerable.Range(1, 8).Select(id =>
                $"<w:endnote w:id=\"{id}\"><w:p><w:r><w:endnoteRef/><w:t>Variant note {id}.</w:t></w:r></w:p></w:endnote>")) + "</w:endnotes>",
            ["word/settings.xml"] = $"<w:settings xmlns:w=\"{w}\"><w:endnotePr><w:pos w:val=\"sectEnd\"/></w:endnotePr></w:settings>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"e\" Type=\"{rel}/endnotes\" Target=\"endnotes.xml\"/><Relationship Id=\"s\" Type=\"{rel}/settings\" Target=\"settings.xml\"/></Relationships>"
        });
    }
}
