using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxAutomaticNoteCasingMatchesOwnedReferenceAndTargetPaint()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int count in new[] { 1, 2 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (lower, upper) in new[] { ("lowerRoman", "upperRoman"), ("lowerLetter", "upperLetter") })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var expected = renderer.RenderPagesAsync(ReadAutomaticNoteCasing(kind, count, mode, upper, true, true), null, default)
                .ToBlockingEnumerable().ToArray();
            var source = ReadAutomaticNoteCasing(kind, count, mode, lower, true, true);
            var reference = source.Paragraphs.SelectMany(p => p.InlineReferences).Single(r => r.Id == "4");
            string raw = lower == "lowerRoman" ? "iv" : "d";
            TestAssert.Equal(raw, reference.DisplayText!);
            TestAssert.True(DocxInlineContent.TryGetAutomaticNoteRun(reference, out var owner) && owner.Text == raw.ToUpperInvariant(),
                "The captured marker owns its resolved glyphs while the counter keeps its formatted label.");
            var target = source.RelatedStories.Single(s => s.Id == "4" && s.Kind == reference.Kind);
            TestAssert.Equal(raw.ToUpperInvariant(), target.Paragraphs[0].Runs.Single(r => r.IsAutomaticNoteStoryMarker).Text);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            // Raw counter labels can change retained subset CIDs independently of paint.
            AssertMetricPaint(expected, actual);
            string painted = string.Concat(actual.SelectMany(CustomMarkPhysicalGlyphs).Select(g => g.Text));
            TestAssert.Equal(count, painted.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(4, actual.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Resolved generated markers cannot suppress neighboring supported content.");
        }
    }

    public static void DocxAutomaticNoteCasingKeepsCounterLabelsAndTargetStylesIndependent()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string format in new[] { "lowerRoman", "lowerLetter" })
        foreach (bool bodyCaps in new[] { false, true })
        foreach (bool targetCaps in new[] { false, true })
        {
            var source = ReadAutomaticNoteCasing(kind, 1, mode, format, bodyCaps, targetCaps);
            string raw = format == "lowerRoman" ? "iv" : "d";
            var reference = source.Paragraphs.SelectMany(p => p.InlineReferences).Single(r => r.Id == "4");
            TestAssert.Equal(raw, reference.DisplayText!);
            TestAssert.True(DocxInlineContent.TryGetAutomaticNoteRun(reference, out var owner), "The exact generated source owner stays admitted.");
            TestAssert.Equal(bodyCaps ? raw.ToUpperInvariant() : raw, owner!.Text);
            var target = source.RelatedStories.Single(s => s.Id == "4" && s.Kind == reference.Kind);
            var targetRun = target.Paragraphs[0].Runs.Single(r => r.IsAutomaticNoteStoryMarker);
            TestAssert.Equal(targetCaps ? raw.ToUpperInvariant() : raw, targetRun.Text);
            var actual = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            string painted = string.Concat(actual.SelectMany(CustomMarkPhysicalGlyphs).Select(g => g.Text));
            TestAssert.Equal(1, painted.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(4, actual.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
            TestAssert.Equal(raw, reference.DisplayText!);
        }
    }

    private static DocxDocument ReadAutomaticNoteCasing(string kind, int count, OoxPdfDocxMarkupMode mode,
        string format, bool bodyCaps, bool targetCaps)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        void SetFormat(XDocument xml)
        {
            foreach (var properties in xml.Descendants(w + kind + "Pr"))
            {
                var number = properties.Element(w + "numFmt");
                if (number is null) properties.Add(new XElement(w + "numFmt", new XAttribute(w + "val", format)));
                else number.SetAttributeValue(w + "val", format);
            }
        }
        void SetCaps(XElement run, bool caps)
        {
            var properties = run.Element(w + "rPr");
            if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
            properties.Elements(w + "caps").Remove();
            properties.Add(new XElement(w + "caps", new XAttribute(w + "val", caps ? "true" : "false")));
        }
        return ReadMutation("CompositionProbes/math-custom-note-" + kind + "-" + count, mode, false, false,
            (xml, _, _) =>
            {
                SetFormat(xml);
                var reference = xml.Descendants(w + kind + "Reference").Single(r => r.Attribute(w + "id")?.Value == "4");
                SetCaps(reference.Parent!, bodyCaps);
            }, packageChange: archive =>
            {
                var settings = archive.GetEntry("word/settings.xml");
                if (settings is not null)
                {
                    XDocument xml; using (var input = settings.Open()) xml = XDocument.Load(input);
                    SetFormat(xml); settings.Delete(); using var output = archive.CreateEntry("word/settings.xml").Open(); xml.Save(output);
                }
                var entry = archive.GetEntry("word/" + kind + "s.xml")!; XDocument notes;
                using (var input = entry.Open()) notes = XDocument.Load(input);
                var target = notes.Root!.Elements(w + kind).Single(n => n.Attribute(w + "id")?.Value == "4");
                SetCaps(target.Descendants(w + kind + "Ref").Single().Parent!, targetCaps);
                entry.Delete(); using var noteOutput = archive.CreateEntry("word/" + kind + "s.xml").Open(); notes.Save(noteOutput);
            });
    }
}
