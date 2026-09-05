using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    internal static string BuildRevisionBalloonPreview(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        return BuildRevisionBalloonPreview(revisions, []);
    }

    internal static string BuildRevisionBalloonPreview(DocxParagraph paragraph)
    {
        return BuildRevisionBalloonPreview(paragraph.Revisions, BuildRevisionTextPreviewLabels());

        IReadOnlyList<string> BuildRevisionTextPreviewLabels()
        {
            var labels = new List<string>(3);
            string? deletedText = BuildRevisionTextPreview(paragraph.Runs, DocxRevisionKind.Deletion);
            if (deletedText is not null)
            {
                labels.Add("Deleted: \"" + deletedText + "\"");
            }

            string? movedFromText = BuildRevisionTextPreview(paragraph.Runs, DocxRevisionKind.MoveFrom);
            if (movedFromText is not null)
            {
                labels.Add("Moved from: \"" + movedFromText + "\"");
            }

            string? movedToText = BuildRevisionTextPreview(paragraph.Runs, DocxRevisionKind.MoveTo);
            if (movedToText is not null)
            {
                labels.Add("Moved to: \"" + movedToText + "\"");
            }

            return labels;
        }
    }

    internal static string BuildRevisionBalloonTitle(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        if (revisions.Count == 0)
        {
            return "Tracked change";
        }

        string[] authors = revisions
            .Select(revision => FirstNonEmpty(revision.Author))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .OrderBy(author => author, StringComparer.Ordinal)
            .ToArray();
        string[] dates = revisions
            .Select(revision => FormatCommentDate(revision.Date))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .OrderBy(date => date, StringComparer.Ordinal)
            .ToArray();
        string owner = authors.Length switch
        {
            0 => "Tracked change",
            1 => authors[0],
            _ => authors.Length.ToString(CultureInfo.InvariantCulture) + " reviewers"
        };
        string? dateLabel = dates.Length switch
        {
            0 => null,
            1 => dates[0],
            _ => dates.Length.ToString(CultureInfo.InvariantCulture) + " dates"
        };
        return dateLabel is null ? owner : owner + " " + dateLabel;
    }

    private static string BuildRevisionBalloonPreview(
        IReadOnlyList<DocxRevisionInfo> revisions,
        IReadOnlyList<string> previewLabels)
    {
        if (revisions.Count == 0)
        {
            return string.Empty;
        }

        string[] labels = BuildRevisionBalloonLabels()
            .Concat(previewLabels)
            .Where(label => label.Length != 0)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(RevisionBalloonLabelPriority)
            .ThenBy(label => label, StringComparer.Ordinal)
            .ToArray();
        return labels.Length == 0 ? "Revision" : string.Join(", ", labels);

        IReadOnlyList<string> BuildRevisionBalloonLabels()
        {
            var revisionLabels = new List<string>(revisions.Count);
            revisionLabels.AddRange(revisions
                .Where(revision => !IsFormattingRevision(revision))
                .Select(RevisionBalloonLabel));
            revisionLabels.AddRange(revisions
                .Where(IsFormattingRevision)
                .GroupBy(ResolveFormattingRevisionFamily)
                .Select(group => BuildFormattingRevisionBalloonLabel(
                    group.Key,
                    group.SelectMany(revision => revision.PropertyElementNames))));
            return revisionLabels;
        }
    }

    private static string? BuildRevisionTextPreview(IEnumerable<DocxTextRun> runs, DocxRevisionKind kind)
    {
        string text = string.Concat(runs
            .Where(run => HasRevisionKind(run, kind))
            .Select(run => run.Text));
        string normalized = NormalizeRevisionPreviewText(text);
        return normalized.Length == 0 || normalized.Length > 42 ? null : normalized;
    }

    private static string NormalizeRevisionPreviewText(string value)
    {
        if (value.Any(c => char.IsControl(c) && !char.IsWhiteSpace(c)))
        {
            return string.Empty;
        }

        return Regex.Replace(value, @"\s+", " ").Trim();
    }

    private static string RevisionBalloonLabel(DocxRevisionInfo revision)
    {
        return revision.Kind switch
        {
            DocxRevisionKind.Insertion => "Inserted text",
            DocxRevisionKind.Deletion => "Deleted text",
            DocxRevisionKind.MoveFrom => "Moved from",
            DocxRevisionKind.MoveTo => "Moved to",
            DocxRevisionKind.RunPropertiesChange => BuildFormattingRevisionBalloonLabel(revision),
            DocxRevisionKind.ParagraphPropertiesChange => BuildFormattingRevisionBalloonLabel(revision),
            DocxRevisionKind.TablePropertiesChange => BuildFormattingRevisionBalloonLabel(revision),
            DocxRevisionKind.TableRowPropertiesChange => BuildFormattingRevisionBalloonLabel(revision),
            DocxRevisionKind.TableCellPropertiesChange => BuildFormattingRevisionBalloonLabel(revision),
            DocxRevisionKind.SectionPropertiesChange => BuildFormattingRevisionBalloonLabel(revision),
            _ => revision.Kind.ToValueString()
        };
    }

    private static int RevisionBalloonLabelPriority(string label)
    {
        if (string.Equals(label, "Inserted text", StringComparison.Ordinal))
        {
            return 10;
        }

        if (string.Equals(label, "Deleted text", StringComparison.Ordinal))
        {
            return 20;
        }

        if (string.Equals(label, "Moved from", StringComparison.Ordinal))
        {
            return 30;
        }

        if (string.Equals(label, "Moved to", StringComparison.Ordinal))
        {
            return 40;
        }

        int? formattingPriority = FormattingRevisionLabelPriority();
        if (formattingPriority is not null)
        {
            return formattingPriority.Value;
        }

        if (label.StartsWith("Deleted: ", StringComparison.Ordinal))
        {
            return 60;
        }

        if (label.StartsWith("Moved from: ", StringComparison.Ordinal))
        {
            return 70;
        }

        if (label.StartsWith("Moved to: ", StringComparison.Ordinal))
        {
            return 80;
        }

        return 100;

        int? FormattingRevisionLabelPriority()
        {
            if (label.StartsWith("Formatted section", StringComparison.Ordinal))
            {
                return 50;
            }

            if (label.StartsWith("Formatted table", StringComparison.Ordinal))
            {
                return 51;
            }

            if (label.StartsWith("Formatted row", StringComparison.Ordinal))
            {
                return 52;
            }

            if (label.StartsWith("Formatted cell", StringComparison.Ordinal))
            {
                return 53;
            }

            if (label.StartsWith("Formatted paragraph", StringComparison.Ordinal))
            {
                return 54;
            }

            if (label.StartsWith("Formatted run", StringComparison.Ordinal))
            {
                return 55;
            }

            if (string.Equals(label, "Formatting change", StringComparison.Ordinal) ||
                label.StartsWith("Formatted ", StringComparison.Ordinal))
            {
                return 56;
            }

            return null;
        }
    }

    private static bool IsFormattingRevision(DocxRevisionInfo revision)
    {
        return revision.Kind is
            DocxRevisionKind.RunPropertiesChange or
            DocxRevisionKind.ParagraphPropertiesChange or
            DocxRevisionKind.TablePropertiesChange or
            DocxRevisionKind.TableRowPropertiesChange or
            DocxRevisionKind.TableCellPropertiesChange or
            DocxRevisionKind.SectionPropertiesChange;
    }

    private static DocxRevisionPropertyFamily ResolveFormattingRevisionFamily(DocxRevisionInfo revision)
    {
        return revision.PropertyChangeFamily ?? revision.Kind switch
        {
            DocxRevisionKind.RunPropertiesChange => DocxRevisionPropertyFamily.Run,
            DocxRevisionKind.ParagraphPropertiesChange => DocxRevisionPropertyFamily.Paragraph,
            DocxRevisionKind.TablePropertiesChange => DocxRevisionPropertyFamily.Table,
            DocxRevisionKind.TableRowPropertiesChange => DocxRevisionPropertyFamily.Row,
            DocxRevisionKind.TableCellPropertiesChange => DocxRevisionPropertyFamily.Cell,
            DocxRevisionKind.SectionPropertiesChange => DocxRevisionPropertyFamily.Section,
            _ => DocxRevisionPropertyFamily.Formatting
        };
    }

    private static string BuildFormattingRevisionBalloonLabel(DocxRevisionInfo revision)
    {
        return BuildFormattingRevisionBalloonLabel(ResolveFormattingRevisionFamily(revision), revision.PropertyElementNames);
    }

    private static string BuildFormattingRevisionBalloonLabel(DocxRevisionPropertyFamily family, IEnumerable<string> propertyElementNames)
    {
        string prefix = family switch
        {
            DocxRevisionPropertyFamily.Run => "Formatted run",
            DocxRevisionPropertyFamily.Paragraph => "Formatted paragraph",
            DocxRevisionPropertyFamily.Table => "Formatted table",
            DocxRevisionPropertyFamily.Row => "Formatted row",
            DocxRevisionPropertyFamily.Cell => "Formatted cell",
            DocxRevisionPropertyFamily.Section => "Formatted section",
            _ => "Formatting change"
        };
        string[] names = propertyElementNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        string[] properties = names
            .OrderBy(name => FormattingRevisionPropertyPriority(name))
            .ThenBy(name => FormatFormattingRevisionPropertyName(family, name), StringComparer.Ordinal)
            .Select(name => FormatFormattingRevisionPropertyName(family, name))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .Take(4)
            .ToArray();
        int hiddenPropertyCount = names
            .Select(name => FormatFormattingRevisionPropertyName(family, name))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .Count() - properties.Length;
        string suffix = hiddenPropertyCount <= 0
            ? string.Empty
            : ", +" + hiddenPropertyCount.ToString(CultureInfo.InvariantCulture) + " more";
        return properties.Length == 0 ? prefix : prefix + ": " + string.Join(", ", properties) + suffix;

        int FormattingRevisionPropertyPriority(string value)
        {
            return family switch
            {
                DocxRevisionPropertyFamily.Run => value switch
                {
                    "rStyle" => 0,
                    "rFonts" => 10,
                    "sz" => 20,
                    "color" => 30,
                    "highlight" => 40,
                    "b" => 50,
                    "i" => 60,
                    "u" => 70,
                    "vertAlign" => 80,
                    "shd" => 90,
                    "strike" => 100,
                    "dstrike" => 110,
                    "caps" => 120,
                    "smallCaps" => 130,
                    "vanish" => 140,
                    "rtl" => 150,
                    "lang" => 160,
                    "eastAsianLayout" => 170,
                    "position" => 180,
                    "kern" => 190,
                    "outline" => 200,
                    "shadow" => 210,
                    "emboss" => 220,
                    "imprint" => 230,
                    "webHidden" => 240,
                    "szCs" => 250,
                    "bCs" => 260,
                    "iCs" => 270,
                    "spacing" => 280,
                    "w" => 290,
                    "bdr" => 300,
                    "effect" => 310,
                    "em" => 320,
                    "fitText" => 330,
                    "snapToGrid" => 340,
                    "noProof" => 350,
                    "specVanish" => 360,
                    "cs" => 370,
                    "oMath" => 380,
                    _ => 500
                },
                DocxRevisionPropertyFamily.Paragraph => value switch
                {
                    "pStyle" => 0,
                    "numPr" => 10,
                    "ilvl" => 20,
                    "numId" => 30,
                    "numberingChange" => 40,
                    "jc" => 50,
                    "ind" => 60,
                    "spacing" => 70,
                    "tabs" => 80,
                    "keepNext" => 90,
                    "keepLines" => 100,
                    "widowControl" => 110,
                    "pageBreakBefore" => 120,
                    "outlineLvl" => 130,
                    "pBdr" => 140,
                    "shd" => 150,
                    "contextualSpacing" => 160,
                    "wordWrap" => 170,
                    "bidi" => 180,
                    "textAlignment" => 190,
                    "framePr" => 200,
                    "suppressLineNumbers" => 210,
                    "adjustRightInd" => 220,
                    "autoSpaceDE" => 230,
                    "autoSpaceDN" => 240,
                    "snapToGrid" => 250,
                    "kinsoku" => 260,
                    "overflowPunct" => 270,
                    "topLinePunct" => 280,
                    "suppressAutoHyphens" => 290,
                    "mirrorIndents" => 300,
                    "suppressOverlap" => 310,
                    "cnfStyle" => 320,
                    "divId" => 330,
                    _ => 500
                },
                DocxRevisionPropertyFamily.Table => value switch
                {
                    "tblStyle" => 0,
                    "tblBorders" => 10,
                    "tblW" => 20,
                    "tblLayout" => 30,
                    "tblInd" => 40,
                    "tblCellMar" => 50,
                    "tblLook" => 60,
                    "jc" => 70,
                    "tblCellSpacing" => 80,
                    "shd" => 90,
                    "bidiVisual" => 100,
                    "tblpPr" => 110,
                    "tblOverlap" => 120,
                    "tblCaption" => 130,
                    "tblDescription" => 140,
                    "tblStyleRowBandSize" => 150,
                    "tblStyleColBandSize" => 160,
                    "cnfStyle" => 170,
                    _ => 500
                },
                DocxRevisionPropertyFamily.Row => value switch
                {
                    "tblHeader" => 0,
                    "trHeight" => 10,
                    "cantSplit" => 20,
                    "jc" => 30,
                    "tblCellSpacing" => 40,
                    "gridBefore" => 50,
                    "gridAfter" => 60,
                    "wBefore" => 70,
                    "wAfter" => 80,
                    "hidden" => 90,
                    "tblPrEx" => 100,
                    "cnfStyle" => 110,
                    "divId" => 120,
                    _ => 500
                },
                DocxRevisionPropertyFamily.Cell => value switch
                {
                    "tcW" => 0,
                    "gridSpan" => 10,
                    "vMerge" => 20,
                    "vAlign" => 30,
                    "tcMar" => 40,
                    "tcBorders" => 50,
                    "shd" => 60,
                    "textDirection" => 70,
                    "tcFitText" => 80,
                    "noWrap" => 90,
                    "hideMark" => 100,
                    "hMerge" => 110,
                    "cellIns" => 120,
                    "cellDel" => 130,
                    "cellMerge" => 140,
                    "cnfStyle" => 150,
                    _ => 500
                },
                DocxRevisionPropertyFamily.Section => value switch
                {
                    "type" => 0,
                    "pgSz" => 10,
                    "pgMar" => 20,
                    "cols" => 30,
                    "headerReference" => 40,
                    "footerReference" => 50,
                    "pgNumType" => 60,
                    "docGrid" => 70,
                    "lnNumType" => 80,
                    "footnotePr" => 90,
                    "endnotePr" => 100,
                    "pgBorders" => 110,
                    "titlePg" => 120,
                    "vAlign" => 130,
                    "paperSrc" => 140,
                    "textDirection" => 150,
                    "bidi" => 160,
                    "rtlGutter" => 170,
                    "printerSettings" => 180,
                    "formProt" => 190,
                    _ => 500
                },
                _ => 500
            };
        }
    }

    private static string FormatFormattingRevisionPropertyName(DocxRevisionPropertyFamily family, string value)
    {
        if (family == DocxRevisionPropertyFamily.Run)
        {
            string? runName = value switch
            {
                "spacing" => "character spacing",
                "w" => "character scale",
                "bdr" => "border",
                "effect" => "text effect",
                "em" => "emphasis mark",
                "fitText" => "fit text",
                "snapToGrid" => "snap to grid",
                "noProof" => "proofing",
                "specVanish" => "special hidden text",
                "cs" => "complex script",
                "oMath" => "math",
                _ => null
            };
            if (runName is not null)
            {
                return runName;
            }
        }

        return value switch
        {
            "rStyle" => "character style",
            "b" => "bold",
            "bCs" => "complex-script bold",
            "i" => "italic",
            "iCs" => "complex-script italic",
            "u" => "underline",
            "sz" => "font size",
            "szCs" => "complex-script font size",
            "rFonts" => "font",
            "color" => "color",
            "highlight" => "highlight",
            "shd" => "shading",
            "strike" => "strike",
            "dstrike" => "double strike",
            "caps" => "all caps",
            "smallCaps" => "small caps",
            "vanish" => "hidden text",
            "vertAlign" => "vertical position",
            "rtl" => "right-to-left",
            "lang" => "language",
            "eastAsianLayout" => "East Asian layout",
            "position" => "position",
            "kern" => "kerning",
            "outline" => "outline",
            "shadow" => "shadow",
            "emboss" => "emboss",
            "imprint" => "engrave",
            "webHidden" => "web hidden",
            "jc" => "alignment",
            "ind" => "indent",
            "spacing" => "spacing",
            "tabs" => "tab stops",
            "numPr" => "numbering",
            "ilvl" => "list level",
            "numId" => "numbering id",
            "numberingChange" => "numbering change",
            "pStyle" => "paragraph style",
            "keepNext" => "keep next",
            "keepLines" => "keep lines",
            "widowControl" => "widow control",
            "pageBreakBefore" => "page break before",
            "outlineLvl" => "outline level",
            "pBdr" => "borders",
            "contextualSpacing" => "contextual spacing",
            "wordWrap" => "word wrap",
            "bidi" => "right-to-left",
            "textAlignment" => "text alignment",
            "framePr" => "frame",
            "suppressLineNumbers" => "line numbering",
            "adjustRightInd" => "adjust right indent",
            "autoSpaceDE" => "East Asian auto spacing",
            "autoSpaceDN" => "number auto spacing",
            "snapToGrid" => "snap to grid",
            "kinsoku" => "kinsoku",
            "overflowPunct" => "overflow punctuation",
            "topLinePunct" => "top-line punctuation",
            "suppressAutoHyphens" => "auto hyphenation",
            "mirrorIndents" => "mirror indents",
            "suppressOverlap" => "overlap suppression",
            "cnfStyle" => "conditional formatting",
            "divId" => "division",
            "tblW" => "table width",
            "tblBorders" => "borders",
            "tblCellMar" => "cell margins",
            "tblCellSpacing" => "cell spacing",
            "tblInd" => "table indent",
            "tblLayout" => "table layout",
            "tblLook" => "table look",
            "bidiVisual" => "right-to-left table",
            "tblpPr" => "floating table position",
            "tblOverlap" => "table overlap",
            "tblCaption" => "table caption",
            "tblDescription" => "table description",
            "tblStyleRowBandSize" => "row band size",
            "tblStyleColBandSize" => "column band size",
            "trHeight" => "row height",
            "cantSplit" => "row split",
            "tblHeader" => "header row",
            "gridBefore" => "grid before",
            "gridAfter" => "grid after",
            "wBefore" => "width before",
            "wAfter" => "width after",
            "hidden" => "hidden row",
            "tblPrEx" => "table property exceptions",
            "tcW" => "cell width",
            "gridSpan" => "grid span",
            "vMerge" => "vertical merge",
            "hMerge" => "horizontal merge",
            "tcBorders" => "cell borders",
            "tcMar" => "cell margins",
            "vAlign" => "vertical alignment",
            "textDirection" => "text direction",
            "tcFitText" => "fit text",
            "noWrap" => "no wrap",
            "hideMark" => "end mark",
            "cellIns" => "inserted cell",
            "cellDel" => "deleted cell",
            "cellMerge" => "cell merge",
            "pgSz" => "page size",
            "pgMar" => "page margins",
            "cols" => "columns",
            "type" => "section type",
            "headerReference" => "header",
            "footerReference" => "footer",
            "pgNumType" => "page numbering",
            "docGrid" => "document grid",
            "lnNumType" => "line numbering",
            "footnotePr" => "footnotes",
            "endnotePr" => "endnotes",
            "pgBorders" => "page borders",
            "titlePg" => "different first page",
            "paperSrc" => "paper source",
            "rtlGutter" => "right-to-left gutter",
            "printerSettings" => "printer settings",
            "formProt" => "form protection",
            _ => HumanizeFormattingRevisionPropertyName()
        };

        string HumanizeFormattingRevisionPropertyName()
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string spaced = Regex.Replace(value.Trim(), "([a-z0-9])([A-Z])", "$1 $2");
            return spaced.Replace(" Pr", " properties", StringComparison.Ordinal).ToLowerInvariant();
        }
    }
}
