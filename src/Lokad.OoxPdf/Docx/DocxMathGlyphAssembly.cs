using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Docx;

internal readonly record struct DocxMathGlyphPartPosition(OpenTypeMathGlyphPart Part, double Offset);

// Font design units only. Root placement supplies the baseline and painting
// keeps one source symbol for the complete assembly.
internal static class DocxMathGlyphAssembly
{
    public static bool TryArrange(OpenTypeMathGlyphAssembly assembly, double targetHeight, CancellationToken token,
        out IReadOnlyList<DocxMathGlyphPartPosition>? positions, out double height)
    {
        positions = null; height = 0d;
        token.ThrowIfCancellationRequested();
        var parts = assembly.Parts;
        int minimum = assembly.MinimumConnectorOverlap;
        if (!double.IsFinite(targetHeight) || targetHeight < 0d || targetHeight > ushort.MaxValue ||
            parts.Count == 0 || parts.Count > ushort.MaxValue) return false;
        var fixedParts = new List<OpenTypeMathGlyphPart>();
        long growth = 0;
        foreach (var part in parts)
        {
            token.ThrowIfCancellationRequested();
            if (part.Glyph == 0 || part.FullAdvance == 0) return false;
            if (!part.IsExtender) fixedParts.Add(part);
            else growth += Math.Max(0, part.FullAdvance - minimum);
        }
        // First try the non-extending pieces alone, as specified by MATH.
        IReadOnlyList<OpenTypeMathGlyphPart> selected = fixedParts;
        if (!TryRange(selected, out _, out double maximum) || maximum < targetHeight)
        {
            if (!TryRange(parts, out _, out maximum)) return false;
            int additional = 0;
            if (maximum < targetHeight)
            {
                if (growth == 0) return false;
                additional = (int)Math.Ceiling((targetHeight - maximum) / growth);
            }
            long count = parts.Count;
            foreach (var part in parts)
            {
                token.ThrowIfCancellationRequested();
                if (part.IsExtender && part.FullAdvance > minimum) count += additional;
            }
            if (count > int.MaxValue) return false;
            var expanded = new List<OpenTypeMathGlyphPart>((int)count);
            foreach (var part in parts)
            {
                // A full-overlap extender adds no height. Repeating it would
                // paint the same piece at the same origin with no visible gain.
                int copies = part.IsExtender && part.FullAdvance > minimum ? additional + 1 : 1;
                for (int copy = 0; copy < copies; copy++)
                { token.ThrowIfCancellationRequested(); expanded.Add(part); }
            }
            selected = expanded;
        }
        if (!TryRange(selected, out double smallest, out double largest) || largest < targetHeight ||
            smallest > ushort.MaxValue) return false;
        double actualHeight = Math.Max(smallest, targetHeight);
        var capacities = new int[selected.Count - 1];
        for (int index = 0; index < capacities.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            capacities[index] = MaximumOverlap(selected[index], selected[index + 1]) - minimum;
        }
        // Distribute growth equally until a connection reaches its minimum
        // overlap, then distribute the remainder over the other connections.
        Array.Sort(capacities);
        token.ThrowIfCancellationRequested();
        double remaining = actualHeight - smallest, reduction = 0d;
        for (int index = 0; index < capacities.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            int active = capacities.Length - index;
            double available = (capacities[index] - reduction) * active;
            if (remaining <= available)
            { reduction += remaining / active; remaining = 0d; break; }
            remaining -= available;
            reduction = capacities[index];
        }
        var result = new DocxMathGlyphPartPosition[selected.Count];
        double offset = 0d;
        for (int index = 0; index < selected.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            result[index] = new(selected[index], offset);
            if (index + 1 < selected.Count)
            {
                int overlap = MaximumOverlap(selected[index], selected[index + 1]);
                offset += selected[index].FullAdvance - overlap + Math.Min(reduction, overlap - minimum);
            }
        }
        height = offset + selected[^1].FullAdvance;
        positions = Array.AsReadOnly(result);
        return true;

        bool TryRange(IReadOnlyList<OpenTypeMathGlyphPart> sequence, out double smallest, out double largest)
        {
            smallest = largest = 0d;
            if (sequence.Count == 0) return false;
            for (int index = 0; index < sequence.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                smallest += sequence[index].FullAdvance;
                largest += sequence[index].FullAdvance;
                if (index == 0) continue;
                int overlap = MaximumOverlap(sequence[index - 1], sequence[index]);
                if (overlap < minimum) return false;
                smallest -= overlap;
                largest -= minimum;
            }
            return true;
        }
    }

    // Connectors can extend beyond a part's advance (DejaVu's root top does).
    // Overlap must also leave nondecreasing origins and an enclosing endpoint.
    private static int MaximumOverlap(OpenTypeMathGlyphPart before, OpenTypeMathGlyphPart after) =>
        Math.Min(Math.Min(before.EndConnectorLength, after.StartConnectorLength), Math.Min(before.FullAdvance, after.FullAdvance));
}
