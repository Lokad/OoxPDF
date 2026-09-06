using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

// Reader pipeline stages (entry: Core.Read; each file owns one stage):
//   1. Package (Core.cs): main-part resolution plus relationship maps.
//   2. Settings plus markup profile (Core.cs): document settings plus markup context.
//   3. Probes (Probes.cs): comment-anchor inventory plus unsupported-feature diagnostics.
//   4. Catalogs (Styles.cs, TableStyles.cs, Numbering.cs, font loading): style, numbering and font sets.
//   5. Sections (Sections.cs): section breaks plus page settings.
//   6. Body (Body.cs dispatch): paragraphs, runs and fields (Paragraphs.cs plus Fields plus Spacing) plus tables (Tables.cs).
//   7. Stories (Stories.cs): headers, footers, related and comment stories.
//   8. Markup predicates (Markup.cs): revision containers, markers and info, consumed inline by body readers.
//   Shared: Records.cs resolved-property bags travel between stages.
//   Assembly: Core.Read tail builds the DocxDocument from stage outputs.
internal sealed partial class DocxReader
{
    private const double WordUntokenedAutoLineSpacingFactor = 1.2d;
    private const double WordSpacingTokenAutoLineSpacingFactor = 1.2d;
    private const double WordDefaultSpacingAfterPoints = 8d;

    private sealed class DocxComplexFieldState
    {
        public DocxComplexFieldState(int sourceRunIndex, int textRunIndex, int textLengthStart, int nestingDepth)
        {
            SourceRunIndex = sourceRunIndex;
            TextRunIndex = textRunIndex;
            TextLengthStart = textLengthStart;
            NestingDepth = nestingDepth;
        }

        public StringBuilder Instruction { get; } = new();
        public int SourceRunIndex { get; }
        public int InstructionSourceRunIndex { get; set; } = -1;
        public int TextRunIndex { get; private set; }
        public int TextLengthStart { get; private set; }
        public int NestingDepth { get; }
        public bool HasSeparate { get; set; }
        public bool InResult { get; set; }
        public bool HasCachedResult { get; set; }
        public bool RendersCachedResult { get; set; }
        public bool PlaceholderEmitted { get; set; }
        public int InstructionRunCount { get; set; }
        public int ResultRunCount { get; set; }

        public void EnsureTextSpan(int textRunIndex, int textLengthStart)
        {
            if (ResultRunCount == 0 && !PlaceholderEmitted && !RendersCachedResult)
            {
                TextRunIndex = textRunIndex;
                TextLengthStart = textLengthStart;
            }
        }
    }

    // Scan-phase projection of complex-field open/close tracking, used only by the
    // unsupported-field pre-scan (HasUnsupportedComplexFields family). The full parse
    // carries richer per-field indices in DocxComplexFieldState; this stays separate
    // because the pre-scan runs before (and independently of) paragraph parsing.
    private sealed record DocxComplexFieldScanState(
        StringBuilder Instruction,
        bool HasSeparate,
        bool InResult,
        bool HasCachedResult);

    private const string MainDocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string NumberingRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    private const string HeaderRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    private const string FooterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    private const string SettingsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    private const string FontTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    private const string ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    private const string CommentsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    private const string CommentsExtendedRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    private const string FootnotesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    private const string EndnotesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    private const string StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    private const string NumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string FontTableContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    private const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string CommentsExtendedContentType = "application/vnd.ms-word.commentsExtended+xml";
    private const string FootnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    private const string EndnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    private const double WordAutomaticParagraphSpacingPoints = 14d;

}
