namespace Lokad.OoxPdf.Docx;

internal readonly record struct DocxTextAdvances(bool PairKerningEnabled, bool UseLogicalSizeForPdfAdvances)
{
    // Resolve a rendering purpose without rewriting source runs or the state
    // consumed by existing note/numbering eligibility. Shifted glyph sizes and
    // automatic marker owners retain their separately qualified purposes.
    internal static DocxTextAdvances Resolve(DocxTextRun run)
    {
        if (run.SourceRunIndex < 0 || run.VerticalAlignment != DocxRunVerticalAlignment.Baseline ||
            run.IsAutomaticNoteReferenceMarker || run.IsAutomaticNoteStoryMarker ||
            !DocxKerningThreshold.TryResolve(run, out bool enabled))
            return new(run.PairKerningEnabled, run.UseLogicalSizeForPdfAdvances);

        return new(run.PairKerningEnabled && enabled, true);
    }
}
