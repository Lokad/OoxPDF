using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private readonly record struct DocxResolvedRunProperties(
        double? FontSize,
        string? ColorHex,
        string? FontFamily,
        bool? Bold,
        bool? Italic,
        bool? ComplexScriptBold,
        bool? ComplexScriptItalic,
        bool? Underline,
        string? UnderlineValue,
        DocxRunFonts Fonts,
        double? CharacterSpacingPoints,
        bool? AllCaps,
        string? VerticalAlignmentValue,
        bool? Strike,
        string? StrikeValue,
        bool? DoubleStrike,
        string? DoubleStrikeValue,
        string? HighlightValue,
        string? ShadingFillHex,
        string? ShadingValue,
        string? ShadingColor,
        bool? SmallCaps,
        string? SmallCapsValue,
        bool? Hidden,
        string? HiddenValue,
        string? UnderlineColorHex)
    {
        public static DocxResolvedRunProperties Empty { get; } = new(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

        public DocxResolvedRunProperties Merge(DocxResolvedRunProperties other)
        {
            return new DocxResolvedRunProperties(
                other.FontSize ?? FontSize,
                other.ColorHex ?? ColorHex,
                other.FontFamily ?? FontFamily,
                other.Bold ?? Bold,
                other.Italic ?? Italic,
                other.ComplexScriptBold ?? ComplexScriptBold,
                other.ComplexScriptItalic ?? ComplexScriptItalic,
                other.Underline ?? Underline,
                other.UnderlineValue ?? UnderlineValue,
                MergeRunFonts(Fonts, other.Fonts),
                other.CharacterSpacingPoints ?? CharacterSpacingPoints,
                other.AllCaps ?? AllCaps,
                other.VerticalAlignmentValue ?? VerticalAlignmentValue,
                other.Strike ?? Strike,
                other.StrikeValue ?? StrikeValue,
                other.DoubleStrike ?? DoubleStrike,
                other.DoubleStrikeValue ?? DoubleStrikeValue,
                other.HighlightValue ?? HighlightValue,
                other.ShadingFillHex ?? ShadingFillHex,
                other.ShadingValue ?? ShadingValue,
                other.ShadingColor ?? ShadingColor,
                other.SmallCaps ?? SmallCaps,
                other.SmallCapsValue ?? SmallCapsValue,
                other.Hidden ?? Hidden,
                other.HiddenValue ?? HiddenValue,
                other.UnderlineColorHex ?? UnderlineColorHex);
        }
    }

    private sealed record DocxCommentAnchorInventory(
        IReadOnlyList<string> PackageAnchorIds,
        IReadOnlyList<string> HiddenAnchorIds);

    private sealed record DocxCommentThreadMetadata(string? ParentParagraphId, bool? IsResolved);
}
