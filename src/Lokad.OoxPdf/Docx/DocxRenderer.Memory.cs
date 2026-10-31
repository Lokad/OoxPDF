namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // One page's paint/source family, including snapshot and child-vector overlap.
    // This is coarse construction capacity, not per-operation work accounting.
    private static IMemoryReservation? ReservePageConstruction(DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floating, IReadOnlyList<DocxFloatingDrawingLayout> staticDrawings,
        CancellationToken token)
    {
        IMemoryScope? owner = OoxConversionBudget.Current?.Memory;
        if (owner is null) return null;
        long capacity = 64L * 1024;
        Items(page.Items);
        Items(page.StaticTextLines); Items(page.StaticInlineImages);
        Items(page.StaticInlineTextBoxes); Items(page.StaticTableRows);
        Drawings(floating); Drawings(staticDrawings);
        foreach (var story in page.PlacedRelatedStories)
        {
            token.ThrowIfCancellationRequested();
            Items(story.TextLines); Items(story.InlineImages); Items(story.TableRows); Drawings(story.FloatingDrawings);
        }
        token.ThrowIfCancellationRequested();
        return OoxMemory.Reserve(owner, capacity);

        void Items(IEnumerable<DocxLayoutItem> items)
        {
            foreach (var item in items)
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 4096);
                switch (item)
                {
                    case DocxTextLineLayout line:
                        capacity = checked(capacity + 512L * line.Text.Length + 1024L * line.Segments.Count);
                        if (line.Math is { } math) Math(math);
                        foreach (var inline in line.InlineMath) Math(inline.Math);
                        break;
                    case DocxEquationLayout equation: Math(equation.Math); break;
                    case DocxInlineImageLayout image: Image(image.Image); break;
                    case DocxInlineTextBoxLayout box:
                        Items(box.TextLines); Items(box.InlineImages); Items(box.TableRows);
                        break;
                    case DocxTableRowLayout row:
                        foreach (var cell in row.Cells)
                        {
                            token.ThrowIfCancellationRequested();
                            capacity = checked(capacity + 4096);
                            Items(cell.TextLines); Items(cell.InlineImages); Items(cell.InlineTextBoxes); Items(cell.NestedRows);
                        }
                        break;
                }
            }
        }
        void Math(DocxPreparedMathParagraph math)
        {
            capacity = checked(capacity + 4096L * math.Box.Glyphs.Count + 1024L * math.Box.Rules.Count);
            foreach (var glyph in math.Box.Glyphs)
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 512L * glyph.AssemblyParts.Count);
            }
        }
        void Image(DocxInlineImage image)
        {
            if (image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) && image.Bytes is { } bytes)
                capacity = checked(capacity + 64L * bytes.Length);
        }
        void Drawings(IReadOnlyList<DocxFloatingDrawingLayout> drawings)
        {
            foreach (var drawing in drawings)
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 4096);
                if (drawing.Drawing.Image is { } image) Image(image);
                if (drawing.TextBoxLayout is { } box)
                {
                    Items(box.TextLines); Items(box.InlineImages); Items(box.TableRows); Drawings(box.FloatingDrawings);
                }
            }
        }
    }
}
