using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMathRolesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent script/matrix argument frames use installed Windows fonts.");
        CheckMathRoles(null, true);
    }

    internal static void CheckMathRoles(IFontResolver? resolver, bool geometry)
    {
        CheckEquationRevisions(resolver, geometry, "math-roles-office.json", splitMathRuns: true);
        foreach (string role in new[] { "script", "matrix" })
        {
            string fixture = "CompositionProbes/math-role-" + role + "-ins";
            var unsupported = ReadMutation(fixture, OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, m) =>
                {
                    var equation = xml.Descendants(m + "oMath").First();
                    if (role == "script") equation.Descendants(m + "sSubSup").Single().Element(m + "e")!.RemoveNodes();
                    else equation.Descendants(m + "r").Remove();
                });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(unsupported, diagnostics.Add, default).Single();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            TestAssert.True(!DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Any(show => fonts.Contains(show.Font)) &&
                diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"),
                "Empty script bases and wholly empty grids retain their explicit fallback: " + role);
        }
    }
}
