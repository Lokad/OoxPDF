using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxPlainElevenPointBody
{
    internal static bool IsEligible(DocxDocument document, DocxLayout layout)
    {
        bool ordinary = document.RelatedStories.Count == 0 && document.Paragraphs.All(p => p.InlineReferences.Count == 0);
        if (!(ordinary ? DocxPlainBodyFrame.IsEligible(document) : DocxPageFootnoteNumbering.IsEligibleDocument(document)) ||
            !DocxPlainFootnoteKerning.HasSingleLineBody(layout) ||
            layout.Pages.Any(p => p.ColumnFrames.Count > 1 || p.StaticTextLines.Count != 0 || p.StaticInlineImages.Count != 0) ||
            document.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true)) return false;
        string? family = document.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Text.Length != 0)?.FontFamily;
        if (family is null) return false;
        foreach (var paragraph in document.Paragraphs)
        {
            // Resolve the complete frame before admitting a new source size;
            // legacy 12pt admissions retain their independent qualifications.
            if (!DocxMinimumBodyLineLeading.IsQualifiedExpansionSpacing(paragraph, allowContextual: true) ||
                !(paragraph.StyleId is not null ? paragraph.StyleResolution.StyleFound :
                    !string.IsNullOrWhiteSpace(paragraph.ContextualSpacingDefaultStyleId)) ||
                paragraph.LineSpacingPoints is { } pitch && pitch is not (>= 16d and <= 24d) ||
                paragraph.LineSpacingPoints is null && paragraph.LineSpacingFactor != 1d) return false;
            foreach (var run in paragraph.Runs.Where(r => r.Text.Length != 0))
            {
                if (!DocxMinimumBodyLineLeading.IsQualifiedParagraphRun(paragraph, run, 11d) ||
                    !string.Equals(run.FontFamily, family, StringComparison.OrdinalIgnoreCase) ||
                    run.KerningMinimumHalfPointsValue is { } value &&
                    (!ushort.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out ushort minimum) || minimum > 32767)) return false;
            }
        }
        foreach (var paragraph in document.RelatedStories.SelectMany(s => s.Paragraphs))
        foreach (var run in paragraph.Runs)
        {
            if (run.KerningMinimumHalfPointsValue is { } value &&
                (!ushort.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out ushort minimum) || minimum > 32767) ||
                run.Text.Any(char.IsLetter) && (!DocxMinimumBodyLineLeading.IsQualifiedParagraphRun(paragraph, run) ||
                    !string.Equals(run.FontFamily, family, StringComparison.OrdinalIgnoreCase))) return false;
        }
        return true;
    }
}
