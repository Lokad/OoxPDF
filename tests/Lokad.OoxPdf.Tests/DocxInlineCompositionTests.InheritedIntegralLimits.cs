using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionInheritedIntegralLimitsKeepWordContainerPlacement()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent integral origins require installed Windows fonts.");
        CheckInheritedIntegralLimits(null);
    }

    internal static void CheckInheritedIntegralLimits(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("inherited-integral-office.json")));
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string fixture = "CompositionProbes/" + id;
            string part = row.GetProperty("Part").GetString()!;
            string location = row.GetProperty("DocumentIntegralLimit").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            string? originalOrigins = null;
            string? originalContent = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation(fixture, mode, transformed, transformed, partOverride: part);
                var equations = Owners(document).SelectMany(p => p.Equations).ToArray();
                var sourceEquation = equations.Single();
                TestAssert.True(sourceEquation.Display, "The source retains standalone display placement.");
                TestAssert.True(sourceEquation.Content.Children.Single() is DocxMathNary { LimitLocationValue: null },
                    "The source inherits document settings rather than carrying a rewritten node placement.");
                TestAssert.Equal(location, document.Settings.MathLimitSettings.IntegralValue);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderBlankPages(document, diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Inherited limits remain supported.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var math = shows.Where(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)
                    && ReferenceEquals(resource.Font, show.Font))).ToArray();
                TestAssert.Equal(row.GetProperty("MathText").GetString(),
                    string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(4, math.Length);
                TestAssert.True(math.All(show => show.X > 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                    "Every physical math origin remains inside its page.");
                string origins = string.Join("|", math.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                originalContent ??= page.Content;
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
                if (row.TryGetProperty("BodyY", out var bodyY))
                {
                    TestAssert.Contains("Composed story.", DocxPageFootnoteNumberingTests.DecodeText(page));
                    if (resolver is null)
                        TestAssert.True(Math.Abs(shows.First(show => show.Text.StartsWith("Composed", StringComparison.Ordinal)).Y - bodyY.GetDouble()) < 1d,
                            "The occupied header retains independently measured body admission.");
                }
                if (resolver is null)
                {
                    foreach (var anchor in row.GetProperty("PhysicalAnchors").EnumerateArray())
                    {
                        var show = math[anchor.GetProperty("Offset").GetInt32()];
                        TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                            Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d &&
                            Math.Abs(show.Size - anchor.GetProperty("Size").GetDouble()) < 0.1d,
                            "Inherited cell/header limits retain the independent one-point Word gate and authored sizes.");
                    }
                    TestAssert.True(Math.Abs(shows.First(show => show.Text.StartsWith("Follow", StringComparison.Ordinal)).Y -
                        row.GetProperty("FollowingY").GetDouble()) < 1d, "Following prose retains the independent Word flow gate.");
                }
                TestAssert.True(Owners(document).SelectMany(p => p.Equations).SequenceEqual(equations, ReferenceEqualityComparer.Instance),
                    "Preparing inherited defaults preserves original source equation identities.");
            }

            var source = ReadMutation(fixture, mode, false, false, partOverride: part);
            var malformed = DocxMathLimitSettings.Read(new XElement("settings", new XElement(m + "mathPr",
                new XElement(m + "intLim", new XAttribute(m + "val", "future")))), default);
            var invalidDocument = source with { Settings = source.Settings with { MathLimitSettings = malformed } };
            var invalidDiagnostics = new List<OoxPdfDiagnostic>();
            var invalidPage = renderer.RenderBlankPages(invalidDocument, invalidDiagnostics.Add, default).Single();
            TestAssert.True(invalidDiagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"),
                "A malformed used document default retains an explicit fallback in every container.");
            TestAssert.True(!invalidPage.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "Unsupported inherited settings do not leave a partial equation resource.");
            TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(invalidPage));

            var explicitSource = ReadMutation(fixture, mode, false, false, (xml, _, mathNamespace) =>
                xml.Descendants(mathNamespace + "naryPr").Single().Add(new XElement(mathNamespace + "limLoc",
                    new XAttribute(mathNamespace + "val", location))), part);
            var explicitDocument = explicitSource with { Settings = explicitSource.Settings with { MathLimitSettings = malformed } };
            var explicitDiagnostics = new List<OoxPdfDiagnostic>();
            var explicitPage = renderer.RenderBlankPages(explicitDocument, explicitDiagnostics.Add, default).Single();
            TestAssert.True(explicitDiagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Explicit placement overrides a malformed default.");
            TestAssert.True(originalContent == explicitPage.Content,
                "Explicit node placement takes precedence and matches inherited geometry, including with malformed defaults.");

            var unusedMalformed = DocxMathLimitSettings.Read(new XElement("settings", new XElement(m + "mathPr",
                new XElement(m + "naryLim", new XAttribute(m + "val", "future")),
                new XElement(m + "intLim", new XAttribute(m + "val", location)))), default);
            var unusedDocument = source with { Settings = source.Settings with { MathLimitSettings = unusedMalformed } };
            var unusedDiagnostics = new List<OoxPdfDiagnostic>();
            var unusedPage = renderer.RenderBlankPages(unusedDocument, unusedDiagnostics.Add, default).Single();
            TestAssert.True(unusedDiagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Only the used default participates in eligibility.");
            TestAssert.True(originalContent == unusedPage.Content, "An unused malformed default cannot suppress a supported integral.");
        }

        static IEnumerable<DocxParagraph> Owners(DocxDocument document) =>
            DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default, true).Concat(
                DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Select(owner => owner.Paragraph));
    }
}
