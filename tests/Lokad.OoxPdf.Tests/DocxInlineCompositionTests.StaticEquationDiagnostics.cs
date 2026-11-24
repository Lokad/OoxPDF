using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxEquationWarningsOmitUnselectedStaticParts()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadStaticVariantPropertySource(story, variant, mode, history: true, enabled: false);
            TestAssert.True(RenderEquationViewDiagnostics(source, mode).All(d =>
                d.Id is not ("DOCX_UNSUPPORTED_EQUATION" or "DOCX_APPROXIMATED_EQUATION")),
                "A deliberately unselected static part cannot warn that visible mathematics was ignored.");
        }
        // Enabled does not mean printed: this one-page document has no even page.
        foreach (string story in new[] { "header", "footer" })
        {
            var source = ReadStaticVariantPropertySource(story, "even", OoxPdfDocxMarkupMode.Final, history: false, enabled: true);
            TestAssert.True(RenderEquationViewDiagnostics(source, OoxPdfDocxMarkupMode.Final)
                .All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Use final page selection, not pre-layout reachability.");
        }
    }

    public static void DocxEquationWarningsKeepActiveSharedStaticPartsAcrossPages()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (bool reverse in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadSharedStaticEquationSource(story, mode, reverse);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            string part = "/word/inline-box-" + story + ".xml";
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == part),
                "A shared part selected on an earlier page retains its genuinely unsupported neighbor warning.");
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION" && d.PartName == part),
                "The supported equation still paints in its actual part.");
            TestAssert.Contains("Before inline", DocxPageFootnoteNumberingTests.DecodeText(pages[0]));
            TestAssert.Contains("Later plain story.", DocxPageFootnoteNumberingTests.DecodeText(pages[1]));
            foreach (var page in pages) Lokad.OoxPdf.Pdf.PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    public static void DocxEquationWarningsRetainUnresolvedStaticBindings()
    {
        var source = ReadStaticVariantPropertySource("header", "first", OoxPdfDocxMarkupMode.Final, history: false, enabled: false);
        var part = source.EquationParts.Single(p => p.PartName == "/word/inline-box-header.xml");
        foreach (var unresolved in new[] { part with { BodyElements = part.BodyElements.ToArray() }, part with { BodyElements = [] } })
        {
            var changed = source with { EquationParts = source.EquationParts.Select(p => ReferenceEquals(p, part) ? unresolved : p).ToArray() };
            TestAssert.True(RenderEquationViewDiagnostics(changed, OoxPdfDocxMarkupMode.Final)
                .Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == part.PartName),
                "A copied or empty source list cannot prove that a warning belongs only to an unprinted part.");
        }
        var alias = part with { PartName = "/word/ambiguous-source.xml" };
        var warning = source.EquationDiagnostics.Single(d => d.PartName == part.PartName);
        var ambiguous = source with { EquationParts = source.EquationParts.Append(alias).ToArray(),
            EquationDiagnostics = source.EquationDiagnostics.Append(warning with { PartName = alias.PartName }).ToArray() };
        var actual = RenderEquationViewDiagnostics(ambiguous, OoxPdfDocxMarkupMode.Final);
        foreach (string name in new[] { part.PartName, alias.PartName })
            TestAssert.True(actual.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == name),
                "Multiple part bindings retain their warnings rather than guessing an owner.");
    }

    public static void DocxStaticEquationDiagnosticsPreserveObserverFailuresAndCancellation()
    {
        var source = ReadSharedStaticEquationSource("header", OoxPdfDocxMarkupMode.Final, reverse: false);
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
            OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        bool observed = false;
        TestAssert.Throws<InvalidOperationException>(() => renderer.RenderPagesAsync(source, d =>
        {
            if (d.Id != "DOCX_UNSUPPORTED_EQUATION") return;
            observed = true; throw new InvalidOperationException("Expected active equation observer failure.");
        }, default).ToBlockingEnumerable().ToArray());
        TestAssert.True(observed, "A genuine unsupported active source still reaches its observer.");
        observed = false;
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, d =>
        {
            if (d.Id != "DOCX_UNSUPPORTED_EQUATION") return;
            observed = true; cancellation.Cancel();
        }, cancellation.Token).ToBlockingEnumerable().ToArray());
        TestAssert.True(observed, "Cancellation after reporting begins still stops the remaining diagnostics and publication.");
    }

    private static DocxDocument ReadSharedStaticEquationSource(string story, OoxPdfDocxMarkupMode mode, bool reverse)
    {
        var source = ReadMutation("InlineBoxMath/inline-box-math-" + story, mode, false, false,
            (xml, _, m) => xml.Descendants(m + "oMath").Single().AddAfterSelf(new XElement(m + "oMath", new XElement(m + "future"))),
            partOverride: "word/inline-box-" + story + ".xml");
        var settings = StaticVariantSettings(source.PageSettings, story, "first", enabled: true);
        var plain = DocxTests.CreateDocxLayoutParagraph("Later plain story.", 12d, 12d);
        var originalBodies = story == "header" ? settings.HeaderBodyElementsByType : settings.FooterBodyElementsByType;
        var body = originalBodies["first"];
        // The same cached part is also referenced by an inactive variant. The
        // last page selects a distinct plain story; earlier selection must win.
        string[] variants = reverse ? ["default", "even", "first"] : ["first", "even", "default"];
        var bodies = variants.ToDictionary(key => key, key => key == "default" ?
            (IReadOnlyList<DocxBodyElement>)[new DocxParagraphElement(plain)] : body, StringComparer.OrdinalIgnoreCase);
        settings = story == "header" ? settings with { HeaderBodyElementsByType = bodies } :
            settings with { FooterBodyElementsByType = bodies };
        source = WithStaticVariantSettings(source, settings);
        return source with { BodyElements = source.BodyElements.Concat<DocxBodyElement>(
            [new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
            new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Second page body.", 12d, 12d))]).ToArray() };
    }
}
