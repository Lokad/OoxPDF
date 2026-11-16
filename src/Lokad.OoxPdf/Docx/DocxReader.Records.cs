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
        public string? KerningMinimumHalfPointsValue { get; init; }
        public static DocxResolvedRunProperties Empty { get; } = new(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

        // A basedOn chain resolves by assignment within its style level. Word
        // composes toggle properties between levels relative to document defaults
        // (MS-OI29500 17.7.3); direct formatting still uses ordinary Merge.
        public DocxResolvedRunProperties MergeStyleLevel(DocxResolvedRunProperties other, DocxResolvedRunProperties defaults)
        {
            return Merge(other) with
            {
                Bold = Toggle(Bold, other.Bold, defaults.Bold),
                Italic = Toggle(Italic, other.Italic, defaults.Italic),
                ComplexScriptBold = Toggle(ComplexScriptBold, other.ComplexScriptBold, defaults.ComplexScriptBold),
                ComplexScriptItalic = Toggle(ComplexScriptItalic, other.ComplexScriptItalic, defaults.ComplexScriptItalic),
                AllCaps = Toggle(AllCaps, other.AllCaps, defaults.AllCaps),
                Strike = Toggle(Strike, other.Strike, defaults.Strike),
                SmallCaps = Toggle(SmallCaps, other.SmallCaps, defaults.SmallCaps),
                Hidden = Toggle(Hidden, other.Hidden, defaults.Hidden)
            };

            static bool? Toggle(bool? current, bool? declared, bool? defaultValue) =>
                declared is { } value ? (current ?? false) ^ value ^ (defaultValue ?? false) : current;
        }

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
                Fonts.Merge(other.Fonts),
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
                other.UnderlineColorHex ?? UnderlineColorHex)
            {
                KerningMinimumHalfPointsValue = other.KerningMinimumHalfPointsValue ?? KerningMinimumHalfPointsValue
            };
        }
    }

    private sealed record DocxCommentAnchorInventory(
        IReadOnlyList<string> PackageAnchorIds,
        IReadOnlyList<string> HiddenAnchorIds);

    private sealed record DocxCommentThreadMetadata(string? ParentParagraphId, bool? IsResolved);
}
