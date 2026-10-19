namespace Lokad.OoxPdf.Docx;

internal static class DocxPlainBodyFrame
{
    internal static bool IsOrdinaryBody(DocxDocument document) =>
        document.RelatedStories.Count == 0 && document.Paragraphs.All(p => p.InlineReferences.Count == 0);

    internal static bool HasSingleLineParagraphs(DocxLayout layout) =>
        layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
            .Where(l => l.SourceParagraph is not null && l.Segments.Any(s => s.Role == DocxTextSegmentRole.Text))
            .GroupBy(l => l.SourceParagraph!, ReferenceEqualityComparer.Instance).All(g => g.Count() == 1);

    internal static bool HasSingleLineLayout(DocxDocument document, DocxLayout layout) =>
        HasSingleLineParagraphs(layout) &&
        !layout.Pages.Any(p => p.ColumnFrames.Count > 1) &&
        !document.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true) &&
        (HasNoStaticStories(document)
            ? !layout.Pages.Any(p => p.StaticTextLines.Count != 0 || p.StaticInlineImages.Count != 0)
            : HasIndependentStaticTextLayout(layout));

    internal static bool IsEligible(DocxDocument document, DocxLayout layout) =>
        HasPlainBody(document) && HasSingleLineLayout(document, layout);

    // A measured text story wholly inside its margin does not own the body's
    // line metrics. Rich, unsupported, intrusive and unmeasured stories keep the
    // previous fallback; annotation lanes retain their scaled frame policy.
    private static bool HasIndependentStaticTextLayout(DocxLayout layout)
    {
        if (layout.Pages.Count == 0 || layout.StaticFloatingDrawings.Count != 0) return false;
        var validatedSources = new HashSet<(DocxPageSettings, string)>();
        for (int index = 0; index < layout.Pages.Count; index++)
        {
            var page = layout.Pages[index];
            if (page.MarkupMarginReservePoints != 0d || page.StaticInlineImages.Count != 0 ||
                page.StaticTableRows.Count != 0 || page.StaticInlineTextBoxes.Count != 0 ||
                !layout.HeaderContentBottomByPage.TryGetValue(index, out double headerBottom) ||
                !layout.FooterContentTopByPage.TryGetValue(index, out double footerTop) ||
                double.IsNaN(headerBottom) || double.IsNaN(footerTop) ||
                headerBottom < page.Height - page.MarginTop || footerTop > page.MarginBottom ||
                page.StaticTextLines.Any(line => line.Story?.Kind is not (DocxStoryKind.Header or DocxStoryKind.Footer) ||
                    line.SourceParagraph is not { } paragraph || !IsPlainStaticText(paragraph))) return false;
            var settings = page.PageSettings;
            string variant = DocxLayoutEngine.GetStaticStoryVariant(settings, index + 1);
            if (validatedSources.Add((settings, variant)))
            {
                // Selection belongs to the same container rule as painting.
                // Unused first/even templates do not own this body's metrics.
                var header = DocxBlockTraversal.GetStaticStoryBodyElements(variant,
                    settings.HeaderBodyElementsByType, settings.HeaderParagraphsByType);
                var footer = DocxBlockTraversal.GetStaticStoryBodyElements(variant,
                    settings.FooterBodyElementsByType, settings.FooterParagraphsByType);
                var headerDrawings = settings.HeaderFloatingDrawingsByType.GetValueOrDefault(variant, []);
                var footerDrawings = settings.FooterFloatingDrawingsByType.GetValueOrDefault(variant, []);
                if (!HasPlainStaticTextSource(header.Concat(footer), headerDrawings.Concat(footerDrawings))) return false;
            }
        }
        return true;
    }

    private static bool HasPlainStaticTextSource(IEnumerable<DocxBodyElement> elements,
        IEnumerable<DocxFloatingDrawing> drawings) =>
        !drawings.Any() && elements.All(element =>
            element is DocxParagraphElement paragraph && IsPlainStaticText(paragraph.Paragraph) &&
            DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph.Paragraph, element.Revisions));

    private static bool IsPlainStaticText(DocxParagraph paragraph) =>
        HasPlainTextProperties(paragraph) && HasAutomaticLineFlow(paragraph) && paragraph.InlineReferences.Count == 0 &&
        // PAGE/NUMPAGES already receive page inputs during static measurement.
        // Reuse their source/placeholder contract; arbitrary cached fields do
        // not gain admission merely because their current result is text.
        paragraph.FieldReferences.All(field => field.Kind.IsDynamic()) &&
        DocxInlineContent.HasSourceFieldResults(paragraph.FieldReferences, paragraph.Runs);

    internal static bool IsEligible(DocxDocument document) =>
        HasPlainBody(document) && (HasNoStaticStories(document) || HasNoSelectableStaticStories(document));

    private static bool HasNoSelectableStaticStories(DocxDocument document)
    {
        // Plain source flow has one final section. Before pagination, require
        // no story references for any possible page variant, using layout's selector.
        // A disabled template cannot change numbering, font or leading ownership.
        if (document.HeaderParagraphs.Count != 0 || document.FooterParagraphs.Count != 0) return false;
        var settings = document.PageSettings;
        var variants = document.HeaderParagraphsByType.Keys.Concat(document.FooterParagraphsByType.Keys)
            .Concat(document.HeaderBodyElementsByType.Keys).Concat(document.FooterBodyElementsByType.Keys)
            .Concat(document.HeaderFloatingDrawingsByType.Keys).Concat(document.FooterFloatingDrawingsByType.Keys)
            .Concat(settings.HeaderParagraphsByType.Keys).Concat(settings.FooterParagraphsByType.Keys)
            .Concat(settings.HeaderBodyElementsByType.Keys).Concat(settings.FooterBodyElementsByType.Keys)
            .Concat(settings.HeaderFloatingDrawingsByType.Keys).Concat(settings.FooterFloatingDrawingsByType.Keys);
        string first = DocxLayoutEngine.GetStaticStoryVariant(settings, 1);
        string even = DocxLayoutEngine.GetStaticStoryVariant(settings, 2);
        string ordinary = DocxLayoutEngine.GetStaticStoryVariant(settings, 3);
        return !variants.Any(variant => string.Equals(variant, first, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(variant, even, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(variant, ordinary, StringComparison.OrdinalIgnoreCase));
    }

    internal static bool HasNoStaticStories(DocxDocument document) =>
        document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0 &&
        document.HeaderParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count == 0 &&
        document.HeaderBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count == 0 &&
        document.HeaderFloatingDrawingsByType.Count == 0 && document.FooterFloatingDrawingsByType.Count == 0;

    internal static bool HasKnownBodyFrame(DocxDocument document) =>
        document.FinalSectionBreak is { Revisions.Count: 0 } &&
        document.Settings.MirrorMargins != true && document.FloatingDrawings.Count == 0;

    internal static bool HasPlainBody(DocxDocument document) =>
        HasKnownBodyFrame(document) && document.BodyElements.Count != 0 && document.BodyElements.All(e =>
            e is DocxParagraphElement paragraph ? DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph.Paragraph, e.Revisions) :
            e is DocxPageBreakElement && e.Revisions.Count == 0) &&
        document.Paragraphs.All(IsPlainParagraph);

    internal static bool HasAutomaticLineFlow(DocxParagraph paragraph) =>
        paragraph.LineSpacingPoints is null && paragraph.Spacing.LineRuleValue is (null or "auto");

    internal static bool HasContentDrivenLineFlow(DocxParagraph paragraph) =>
        HasAutomaticLineFlow(paragraph) ||
        paragraph.LineSpacingPoints is { } minimum && double.IsFinite(minimum) && minimum >= 0d &&
        string.Equals(paragraph.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase);

    internal static bool HasFixedLineFlow(DocxParagraph paragraph) =>
        paragraph.LineSpacingPoints is { } height && double.IsFinite(height) && height > 0d &&
        string.Equals(paragraph.Spacing.LineRuleValue, "exact", StringComparison.OrdinalIgnoreCase);

    internal static bool IsPlainParagraph(DocxParagraph p) =>
        p.Alignment == DocxTextAlignment.Left && HasPlainTextContent(p);

    // Bookmarks are non-rendering source metadata. Navigation validates their
    // anchor ownership separately; they do not change the plain flow frame.
    internal static bool HasPlainTextContent(DocxParagraph p) =>
        p.FieldReferences.Count == 0 && HasPlainTextProperties(p);

    private static bool HasPlainTextProperties(DocxParagraph p) =>
        p.Images.Count == 0 && p.InlineTextBoxes.Count == 0 && p.Equations.Count == 0 &&
        p.CommentRanges.Count == 0 && DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(p, p.Revisions) &&
        p.RevisionRanges.Count == 0 && !p.HasDeletedParagraphMark && p.DeletedText.Length == 0 &&
        p.ListLabel is null && p.Hyperlinks.Count == 0 &&
        p.Runs.All(r => r.Revision is null && r.Revisions.Count == 0 &&
            DocxScriptClassifier.IsSimpleLineText(r.Text));
}
