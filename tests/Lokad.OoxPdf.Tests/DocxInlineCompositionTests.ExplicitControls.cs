using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionExplicitControlsKeepIndependentWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent authored controls use installed Windows fonts.");
        CheckEquationRevisions(null, false, "explicit-math-controls-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionExplicitControlsKeepActualOwnersAndMalformedFallbacks()
    {
        foreach (string role in new[] { "script", "fraction" })
        foreach (string paint in new[] { "empty", "painted" })
        foreach (int size in new[] { 11, 12 })
        {
            string fixture = $"CompositionProbes/explicit-control-{role}-{paint}-{size}";
            var equation = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false).Paragraphs.First().Equations.Single();
            bool supported = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var controls = DocxMathLayout.GetControlSources(equation, default);
            TestAssert.True(supported && controls.Count == 2 && controls[1].ControlText?.FontSize == size &&
                controls[1].ControlFontFamily == "Cambria Math" && ((DocxMathNode)controls[1]).PropertiesXml.Count == 1 &&
                DocxMathLayout.GetMetricOnlyControls(equation, runs, default).Count == (paint == "empty" ? 1 : 0),
                "The parsed child owns its actual authored metrics while raw properties and painted sources remain distinct.");
            if (paint != "empty" || size != 11) continue;
            foreach (string malformed in new[] { "unknown", "duplicate-control", "duplicate-owner", "attribute" })
            {
                var invalid = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
                {
                    var child = xml.Descendants(m + (role == "script" ? "sSubSup" : "f")).First();
                    var property = child.Elements().First();
                    if (malformed == "unknown") property.Add(new XElement(m + "future"));
                    else if (malformed == "duplicate-control") property.Add(new XElement(property.Element(m + "ctrlPr")!));
                    else if (malformed == "duplicate-owner") child.AddFirst(new XElement(property));
                    else property.SetAttributeValue("future", "1");
                }).Paragraphs.First().Equations.Single();
                TestAssert.True(!DocxMathLayout.TryGetRuns(invalid, default, out _),
                    "Unknown or malformed controls retain the complete equation fallback.");
            }
        }
    }

    public static void DocxInlineCompositionExplicitControlsKeepEmptyPropertyPaintNeutrality()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native control neutrality uses installed Windows fonts.");
        CheckEmptyControlPropertyNeutrality(null);
    }

    internal static void CheckEmptyControlPropertyNeutrality(IFontResolver? resolver)
    {
        foreach (string role in new[] { "script", "fraction" })
        foreach (string paint in new[] { "empty", "painted" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string? content = null;
            foreach (string variant in new[] { "absent", "empty-owner", "empty-control" })
            {
                var document = ReadMutation($"CompositionProbes/explicit-control-{role}-{paint}-11", mode, false, false, (xml, w, m) =>
                {
                    var child = xml.Descendants(m + (role == "script" ? "sSubSup" : "f")).First();
                    child.Elements().First().Remove();
                    if (variant != "absent") child.AddFirst(new XElement(m + (child.Name.LocalName + "Pr"),
                        variant == "empty-control" ? new XElement(m + "ctrlPr") : null));
                });
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var page = new DocxRenderer(resolver, mode, geometry).RenderBlankPages(document, null, default).Single();
                content ??= page.Content;
                TestAssert.Equal(content, page.Content);
            }
        }
    }
}
