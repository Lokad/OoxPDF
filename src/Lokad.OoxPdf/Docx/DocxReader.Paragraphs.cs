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
    private static DocxParagraph? ReadParagraph(
        XElement paragraph,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle? tableCellStyle,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        DocxRevisionInfo? inheritedRevision,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        bool retainBalloonParagraphMark = false,
        bool isContinuationFragment = false,
        IReadOnlyList<DocxBookmarkAnchor>? leadingBookmarks = null,
        DocxCellBookmarkScope? cellBookmarks = null,
        List<DocxComplexFieldState>? flowFields = null,
        bool continuesAfterBreak = false,
        bool retainEmptyColumnMark = false,
        bool retainUnmatchedCommentStarts = false,
        IReadOnlyList<XElement>? leadingInlineCommentBoundaries = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        string? authoredParagraphStyleId = ReadParagraphStyleId(paragraphProperties);
        string? paragraphStyleId = ReadEffectiveParagraphStyleId(paragraphProperties, styles);
        DocxResolvedParagraphProperties directParagraph = ReadParagraphProperties(paragraphProperties);
        DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
            directParagraph,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        var sourceNumbering = ResolveParagraphNumbering(paragraphProperties, resolvedParagraph, paragraphStyleId, styles, numbering, cancellationToken);
        DocxParagraphStyleResolution styleResolution = CreateParagraphStyleResolution(
            paragraphProperties,
            authoredParagraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        var runs = new List<DocxTextRun>();
        List<DocxNonRenderingRunOwner>? nonRenderingRunOwners = null;
        DocxTextRun? unmarkedListTextStyle = null;
        var images = new List<DocxInlineImage>();
        var equations = new List<DocxEquation>();
        List<DocxInlineItem>? inlineItems = leadingBookmarks is null ? null : [];
        var inlineTextBoxes = new List<DocxInlineTextBox>();
        int hiddenDrawingEquationSourceCount = 0;
        var inlineReferences = new List<DocxInlineReference>();
        var commentRanges = new List<DocxCommentRange>();
        var openCommentRanges = new List<DocxCommentRangeStart>();
        var revisionRanges = new List<DocxRevisionRange>();
        var openRevisionRanges = new List<DocxRevisionRangeStart>();
        var fieldReferences = new List<DocxFieldReference>();
        var hyperlinkSpans = new List<DocxHyperlinkSpan>();
        List<DocxBookmarkAnchor> bookmarkAnchors = leadingBookmarks is null ? [] : [.. leadingBookmarks];
        var paragraphRevisions = new List<DocxRevisionInfo>();
        var deletedText = new StringBuilder();
        AddRevision(paragraphRevisions, inheritedRevision);
        IReadOnlyList<DocxRevisionInfo> propertyRevisions = ReadPropertyChangeRevisions(paragraphProperties);
        AddRevisions(paragraphRevisions, propertyRevisions);
        XElement? paragraphMarkRunProperties = paragraphProperties?.Element(WordprocessingNamespace + "rPr");
        // Resolve the mark once, then publish the same typed revision objects
        // through its actual source owner and aggregate inspection metadata.
        DocxResolvedRunProperties paragraphMarkRun = ResolveRunProperties(
            paragraphMarkRunProperties, paragraphStyleId, characterStyleId: null, styles, tableCellStyle?.Run,
            out DocxRunStyleResolution paragraphMarkStyleResolution);
        IReadOnlyList<DocxRevisionInfo> paragraphMarkRevisions = ReadRunPropertyChangeRevisions(
            paragraphMarkRunProperties, paragraphStyleId, null, styles, tableCellStyle?.Run, paragraphMarkRun);
        DocxRevisionInfo? numberingMarkInsertion = null;
        bool knownNumberingMark = inheritedRevision is null &&
            TryGetNumberingMarkInsertion(paragraphMarkRevisions, cancellationToken, out numberingMarkInsertion);
        bool? originalNumberingIncluded = knownNumberingMark ? numberingMarkInsertion is null : null;
        bool hasDeletedParagraphMark = paragraphMarkRevisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion && revision.SourceElement == "del");
        AddRevisions(paragraphRevisions, paragraphMarkRevisions);
        var complexFieldStack = flowFields ?? new List<DocxComplexFieldState>();
        bool startsInsideComputedField = isContinuationFragment && complexFieldStack.Any(field => field.PlaceholderEmitted);
        int sourceRunIndex = 0;
        long contentSourcePosition = 0;
        HashSet<DocxRevisionInfo>? nonPrintingRangeMarkers = null;
        if (leadingInlineCommentBoundaries is not null)
        {
            foreach (XElement marker in leadingInlineCommentBoundaries)
                AddInlineContainerChild(marker, inheritedRevision);
        }
        foreach (XElement child in paragraph.Elements())
        {
            AddInlineContainerChild(child, inheritedRevision);
        }
        // Normalize only after traversal: an endpoint can close a range outside
        // its immediate revision wrapper. Keep that wrapper's source indices stable.
        if (nonPrintingRangeMarkers is not null)
            paragraphRevisions.RemoveAll(nonPrintingRangeMarkers.Contains);

        bool AddInlineContainerChild(XElement child, DocxRevisionInfo? revision)
            => AddSourceContainerChild(child, revision, removeDeletedText: false);

        // Source ranges do not own revised payload. Unknown children retain a
        // conservative owner so they cannot silently relax layout eligibility.
        bool AddSourceContainerChild(XElement child, DocxRevisionInfo? revision, bool removeDeletedText)
        {
            long before = contentSourcePosition;
            bool hasSource = ParseSourceContainerChild(child, revision, removeDeletedText);
            // A container does not count its already visited children twice.
            // Excluded or unknown owners can still claim source without text runs.
            if (hasSource && contentSourcePosition == before) contentSourcePosition++;
            return hasSource;
        }

        bool ParseSourceContainerChild(XElement child, DocxRevisionInfo? revision, bool removeDeletedText)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (child.Name == MathNamespace + "oMath" || child.Name == MathNamespace + "oMathPara")
            {
                AddEquation(child, revision, removeDeletedText);
                return true;
            }
            else if (removeDeletedText ? DocxSourceXml.IsTransparentInlineContainer(child) :
                DocxSourceXml.IsInlineContentContainer(child))
            {
                return AddVisibleRunContainer(child, revision, removeDeletedText);
            }
            // An excluded deletion retains only supported math structure.
            // Metadata does not change that owner or restore its other content.
            else if (removeDeletedText)
            {
                return false;
            }
            else if (DocxSourceXml.IsNonPrintingMarker(child))
            {
                return false;
            }
            else if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, revision);
                return true;
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                var anchor = ReadBookmarkAnchor(child, sourceRunIndex, runs.Count,
                    RelativeRunTextOffset(runs, sourceRunIndex), GetInlineItemCount());
                if (cellBookmarks is { } scope &&
                    ReadTableColumnBookmark(child, anchor) is { } bookmark && bookmark.First >= scope.EndColumn &&
                    OoxXml.NormalizeDecimalId(anchor.Id) is { Length: > 0 } && !string.IsNullOrEmpty(anchor.Name))
                {
                    scope.Defer(bookmark with { Anchor = ReadBookmarkAnchor(child, 0, 0, 0, 0) });
                }
                else
                {
                    bookmarkAnchors.Add(anchor);
                }
                return false;
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkEnd")
            {
                if (cellBookmarks is { } scope &&
                    OoxXml.NormalizeDecimalId((string?)child.Attribute(WordprocessingNamespace + "id")) is { } id)
                    scope.RecordEnd(id);
                return false;
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                return AddSimpleField(child, revision);
            }
            else if (DocxSourceXml.IsRevisionContainer(child))
            {
                return AddRevisionRunContainer(child, inheritedRevision: revision);
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                return AddHyperlinkContainer(child, revision, AddInlineContainerChild, GetInlineItemCount, ref sourceRunIndex, runs, hyperlinkSpans, relationships);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeStart")
            {
                EnsureInlineContent();
                AddCommentRangeStart(child, openCommentRanges, sourceRunIndex, runs, inlineItems!.Count);
                return false;
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                EnsureInlineContent();
                AddCommentRangeEnd(child, openCommentRanges, commentRanges, sourceRunIndex, runs, inlineItems!.Count);
                return false;
            }
            else if (IsRevisionMarkerElement(child))
            {
                AddRevisionMarker(child, paragraphRevisions, openRevisionRanges, revisionRanges,
                    sourceRunIndex, runs, contentSourcePosition, ref nonPrintingRangeMarkers);
                if (!child.HasElements && string.IsNullOrWhiteSpace(child.Value)) return false;
            }
            return true;
        }

        bool AddVisibleRunContainer(XElement container, DocxRevisionInfo? revision, bool removeDeletedText = false)
        {
            bool hasSource = false;
            foreach (XElement containerChild in DocxSourceXml.StoredChildren(container))
            {
                hasSource |= AddSourceContainerChild(containerChild, revision, removeDeletedText);
            }
            return hasSource;
        }

        bool AddRevisionRunContainer(XElement container, DocxRevisionInfo? inheritedRevision)
        {
            if (!IsIncludedRevisionContainer(container, markupMode))
            {
                // Final/Simple views retain hidden deletions for Office autofit. The
                // clean Original view rejects insertions/move-to content, so those
                // excluded runs must not influence its column widths.
                if (markupMode != OoxPdfDocxMarkupMode.Original)
                {
                    foreach (XElement run in container.Descendants(WordprocessingNamespace + "r"))
                    {
                        deletedText.Append(ReadRunText(run));
                    }
                }

                // Final suppresses deleted mathematical text while supported
                // literal argument frames and structural paint remain.
                if (RetainsDeletedMathStructure(container, markupMode))
                    AddVisibleRunContainer(container, revision: null, removeDeletedText: true);
                return container.HasElements;
            }

            DocxRevisionInfo? revision = CreateRevisionInfo(container) ?? inheritedRevision;
            int revisionIndex = paragraphRevisions.Count;
            if (revision is not null)
            {
                paragraphRevisions.Add(revision);
            }

            bool hasSource = AddVisibleRunContainer(container, revision);
            // Preserve source-order aggregation for actual owners. Annotation-
            // only scopes keep their anchors without governing neighboring ink.
            if (!hasSource && revision is not null) paragraphRevisions.RemoveAt(revisionIndex);
            return hasSource;
        }

        bool instructionContinues = continuesAfterBreak && complexFieldStack.Any(field => !field.InResult);
        FinalizeOpenComplexFields(flowFields is not null);

        bool keepsListMarker = !isContinuationFragment && sourceNumbering is { } numberingSource &&
            !string.IsNullOrWhiteSpace(numberingSource.Definition.Text);
        if ((startsInsideComputedField || instructionContinues) && !retainEmptyColumnMark && !keepsListMarker &&
            runs.Count == 0 && images.Count == 0 && equations.Count == 0 && inlineTextBoxes.Count == 0 &&
            hiddenDrawingEquationSourceCount == 0 && deletedText.Length == 0 &&
            bookmarkAnchors.Count == 0 && hyperlinkSpans.Count == 0 && fieldReferences.Count == 0 &&
            commentRanges.Count == 0 && openCommentRanges.Count == 0 && inlineReferences.Count == 0 &&
            revisionRanges.Count == 0 && openRevisionRanges.Count == 0 && paragraphRevisions.Count == 0)
        {
            // Field code and replaced cache text do not supply a line of
            // content on either side of an explicit break.
            return null;
        }

        // Pilcrow (paragraph-mark) size through the style cascade, excluding direct run
        // formatting. Row-end and break-spill spaces resolve at pilcrow size in Word
        // (spill probes 2026-09-28), so every paragraph carries it even when runs exist.
        if (runs.Count == 0 && images.Count == 0 && inlineTextBoxes.Count == 0)
        {
            AddResolvedTextRun(
                runs,
                string.Empty,
                paragraphMarkRun,
                paragraphMarkStyleResolution,
                complexScript: false,
                sourceRunIndex: -1,
                sourceTextOffsetInRun: 0,
                revision: inheritedRevision,
                revisions: MergeRevisionLists(inheritedRevision, paragraphMarkRevisions));
        }

        // Carry unmatched starts through the same source-offset adjustments as
        // local endpoints. The body reader consumes them before layout.
        if (retainUnmatchedCommentStarts)
        {
            foreach (DocxCommentRangeStart start in openCommentRanges)
            {
                cancellationToken.ThrowIfCancellationRequested();
                commentRanges.Add(new DocxCommentRange(start.Id, start.SourceRunIndex, start.TextOffset, null, null, null, null)
                {
                    StartInlineItemIndex = start.InlineItemIndex
                });
            }
        }

        foreach (DocxRevisionRangeStart openRange in openRevisionRanges)
        {
            revisionRanges.Add(new DocxRevisionRange(
                openRange.Kind,
                openRange.Id,
                openRange.Name,
                openRange.Author,
                openRange.Date,
                openRange.SourceRunIndex,
                openRange.TextOffset,
                EndSourceRunIndex: null,
                EndTextOffset: null));
        }

        double paragraphFontSize = runs.Count == 0 ? DocxDefaults.FontSizePoints : runs.Max(run => run.FontSize);
        var lineSpacing = ResolveLineSpacing(resolvedParagraph);
        double lineSpacingFactor = lineSpacing.Factor ??
            (styles.HasParagraphDefaultOwner ? 1d : WordImplicitNormalAutoLineSpacingFactor);
        double paragraphLineHeight = lineSpacing.Points ?? paragraphFontSize * lineSpacingFactor;
        // A page/column split creates placement fragments, not additional list
        // items. Only the first parsed fragment advances either numbering timeline.
        DocxListLabel? listLabel = isContinuationFragment ? null : ResolveParagraphListLabel();

        DocxTextRun? firstCommentRun = retainBalloonParagraphMark
            ? runs.FirstOrDefault(run => !string.IsNullOrWhiteSpace(run.Text)) : null;
        bool retainMixedMark = retainBalloonParagraphMark &&
            runs.Count(run => !string.IsNullOrWhiteSpace(run.Text)) is 2 or 3 &&
            runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)).All(run =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return DocxBalloonParagraphPolicy.IsPlainBalloonParagraphRun(run);
            });
        bool retainCommentMark = retainBalloonParagraphMark && firstCommentRun is not null &&
            images.Count == 0 && inlineTextBoxes.Count == 0 && fieldReferences.Count == 0 && paragraphRevisions.Count == 0 &&
            runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)).All(run =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return !run.Bold && !run.Italic && !run.Underline && run.CharacterSpacingPoints == 0d && run.FieldKind is null &&
                    (retainMixedMark ||
                        (string.Equals(run.FontFamily, firstCommentRun.FontFamily, StringComparison.OrdinalIgnoreCase) &&
                            Equals(run.EffectiveProperties.Fonts, firstCommentRun.EffectiveProperties.Fonts)));
            });
        DocxTextRun? retainedMark = retainCommentMark || inlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment) ||
            paragraphMarkRevisions.Count != 0
            ? CreateParagraphMark(retainCommentMark ? " " : string.Empty)
            : null;
        if (retainCommentMark && retainedMark is { } commentMark &&
            (commentMark.Bold || commentMark.Italic || commentMark.Underline || commentMark.CharacterSpacingPoints != 0d))
        {
            retainedMark = null;
        }

        if (listLabel is not null && retainedMark is not null && inheritedRevision is null &&
            numberingMarkInsertion is { } markInsertion)
        {
            // The actual numbered mark owns prefix ink. Authored paragraph
            // underline is excluded; explicit level underline remains in Final.
            // Review decoration overrides level decoration, as it does in Word.
            bool review = DocxMarkupContext.FromMode(markupMode).AppliesInlineRevisionStyle;
            var markInk = ApplyMarkupRevisionStyle(paragraphMarkRun with
            {
                Underline = false,
                UnderlineValue = null,
                UnderlineColorHex = null
            }, markInsertion, markupMode);
            listLabel = listLabel with
            {
                DecorationSource = retainedMark with
                {
                    Revision = markInsertion,
                    Revisions = paragraphMarkRevisions,
                    ColorHex = markInk.ColorHex,
                    Underline = markInk.Underline ?? false,
                    UnderlineValue = markInk.UnderlineValue,
                    UnderlineColorHex = markInk.UnderlineColorHex
                },
                Style = review ? listLabel.Style with
                {
                    Underline = true,
                    UnderlineValue = "single",
                    UnderlineColorHex = null
                } : listLabel.Style
            };
        }

        var result = new DocxParagraph(
            runs,
            images,
            authoredParagraphStyleId,
            resolvedParagraph.Alignment ?? DocxTextAlignment.Left,
            resolvedParagraph.AlignmentValue,
            ResolveParagraphSpacingPoints(resolvedParagraph, paragraphLineHeight, before: true),
            ResolveParagraphSpacingPoints(resolvedParagraph, paragraphLineHeight, before: false),
            lineSpacingFactor,
            lineSpacing.Points,
            resolvedParagraph.Spacing,
            resolvedParagraph.KeepRules,
            listLabel)
        {
            Shading = resolvedParagraph.Shading,
            Indent = resolvedParagraph.Indent,
            DirectIndent = directParagraph.Indent,
            TabStops = resolvedParagraph.TabStops,
            SnapToGrid = resolvedParagraph.SnapToGrid,
            SnapToGridValue = resolvedParagraph.SnapToGridValue,
            WordWrap = resolvedParagraph.WordWrap,
            WordWrapValue = resolvedParagraph.WordWrapValue,
            StyleResolution = styleResolution,
            ContextualSpacingDefaultStyleId = paragraphProperties?.Element(WordprocessingNamespace + "pStyle") is null
                ? styles.ContextualSpacingDefaultStyleId : null,
            InlineTextBoxes = inlineTextBoxes,
            DrawingOwners = paragraph.Annotation<DocxParagraphDrawingOwner>() is { } drawingOwner ? [drawingOwner] : [],
            HiddenDrawingEquationSourceCount = hiddenDrawingEquationSourceCount,
            Equations = equations,
            // Page fields retain their canonical source placeholders. Layout
            // eligibility separately requires the actual page inputs.
            // Keep source order independently of revision eligibility. A later
            // paragraph join can consume a mark; layout still checks its owners.
            InlineContent = inlineItems is not null && DocxInlineContent.HasSourceFieldResults(fieldReferences, runs) && DocxInlineContent.HasSupportedReferences(inlineReferences, runs, cancellationToken)
                ? inlineItems : null,
            InlineReferences = inlineReferences,
            CommentRanges = commentRanges,
            RevisionRanges = revisionRanges,
            FieldReferences = fieldReferences,
            Hyperlinks = hyperlinkSpans,
            BookmarkAnchors = bookmarkAnchors,
            Revisions = paragraphRevisions,
            PropertyRevision = propertyRevisions is [{ Kind: DocxRevisionKind.ParagraphPropertiesChange } propertyRevision]
                ? propertyRevision : null,
            NonRenderingRunOwners = nonRenderingRunOwners is null ? [] : nonRenderingRunOwners,
            HasDeletedParagraphMark = hasDeletedParagraphMark,
            DeletedText = deletedText.ToString(),
            ParagraphMarkFontSize = paragraphMarkRun.FontSize ?? DocxDefaults.UnstyledRunFontSizePoints,
            ParagraphMarkRun = retainedMark
        };

        // Retain measurement-only marks through the same source rule used by
        // font planning and static placement. Mixed equation/object owners keep
        // their established mark and font-resource contract.
        return result.ParagraphMarkRun is null && DocxPlainBodyFrame.IsPlainPictureOnlyParagraph(result)
            ? result with { ParagraphMarkRun = CreateParagraphMark(string.Empty) } : result;

        DocxTextRun CreateParagraphMark(string text) => CreateResolvedTextRun(text, paragraphMarkRun,
            paragraphMarkStyleResolution,
            false, -1, 0, inheritedRevision, MergeRevisionLists(inheritedRevision, paragraphMarkRevisions));

        void EnsureInlineContent()
        {
            if (inlineItems is null)
            {
                // Capture source boundaries as they are read. Metadata must not
                // depend on later math or links initializing the ordered view.
                inlineItems = DocxInlineContent.FromPrefix(runs, images, cancellationToken, inlineTextBoxes);
            }
        }

        void PublishInlineRunItems(int firstRun, int firstImage, int? firstBox = null)
        {
            // Drawings retain the same ordered ownership as math and text.
            // Capture the preceding source once, then append this run's text,
            // pictures and boxes together without duplicating current items.
            if (inlineItems is null && (images.Count > firstImage ||
                firstBox is int boxStart && inlineTextBoxes.Count > boxStart))
                inlineItems = DocxInlineContent.FromPrefix(runs.Take(firstRun).ToArray(), images.Take(firstImage).ToArray(), cancellationToken,
                    inlineTextBoxes.Take(firstBox ?? inlineTextBoxes.Count).ToArray());
            if (inlineItems is not null)
                DocxInlineContent.AppendRun(inlineItems, runs.Skip(firstRun), images.Skip(firstImage), cancellationToken,
                    inlineTextBoxes.Skip(firstBox ?? inlineTextBoxes.Count));
        }

        int GetInlineItemCount()
        {
            EnsureInlineContent();
            return inlineItems!.Count;
        }

        void AddEquation(XElement container, DocxRevisionInfo? revision, bool removeDeletedText = false)
        {
            EnsureInlineContent();
            bool display = container.Name == MathNamespace + "oMathPara";
            string? alignment = display
                ? (string?)container.Element(MathNamespace + "oMathParaPr")?.Element(MathNamespace + "jc")?.Attribute(MathNamespace + "val")
                : null;
            Dictionary<XElement, DocxRevisionInfo>? mathRunRevisions = null;
            bool hasMathRunRevisions = false;
            DocxRevisionInfo? equationRevision = revision;
            foreach (var source in EnumerateMathParagraphSources(container, revision, removeDeletedText, markupMode, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                XElement equation = source.Element;
                if (!source.IsEquation)
                {
                    AddSourceContainerChild(equation, source.Revision, source.RemoveDeletedText);
                    continue;
                }
                equationRevision = source.Revision;
                hasMathRunRevisions = false;
                // Trailing text-wrapping breaks belong to the paragraph lines.
                // The preceding operands keep their original equation owner.
                var firstTrailingBreak = display && !source.RemoveDeletedText
                    ? FindFirstTrailingDisplayMathBreak(equation, equationRevision?.Kind == DocxRevisionKind.Deletion,
                        cancellationToken) : null;
                var content = DocxMathParser.Read(equation, ReadMathRun, cancellationToken,
                    wrapper => IsIncludedRevisionContainer(wrapper, markupMode), preserveDeletedStructure: markupMode == OoxPdfDocxMarkupMode.Final,
                    readBookmark: source.RemoveDeletedText ? null : bookmark =>
                        bookmarkAnchors.Add(new DocxBookmarkAnchor(bookmark.Id, bookmark.Name, sourceRunIndex, runs.Count,
                            RelativeRunTextOffset(runs, sourceRunIndex))
                            { InlineItemIndex = inlineItems!.Count, MathSource = bookmark }),
                    readCommentBoundary: source.RemoveDeletedText ? null : boundary =>
                    {
                        // Keep the equation item as the frame owner, even when
                        // the marker itself belongs to a fraction argument.
                        if (boundary.Name.LocalName == "commentRangeStart")
                            AddCommentRangeStart(boundary, openCommentRanges, sourceRunIndex, runs, inlineItems!.Count);
                        else
                            AddCommentRangeEnd(boundary, openCommentRanges, commentRanges, sourceRunIndex, runs, inlineItems!.Count + 1);
                    }, readHyperlink: source.RemoveDeletedText ? null : hyperlink =>
                        AddHyperlinkSpan(hyperlink, sourceRunIndex, 0, runs.Count, 0,
                            inlineItems!.Count, inlineItems.Count + 1, hyperlinkSpans, relationships, mathContent: true),
                    authoredStory: paragraph.Annotation<DocxSourceXml.StoryOrigin>()?.Story,
                    equationHasDeletionOwner: source.RemoveDeletedText || equationRevision?.Kind == DocxRevisionKind.Deletion,
                    trailingBreaks: firstTrailingBreak);
                if (source.RemoveDeletedText)
                {
                    var deletedEquation = new DocxEquation(content, runs.Count, sourceRunIndex, display, null, null);
                    if (DocxMathLayout.RemoveDeletedLiteralMathText(deletedEquation, cancellationToken) is not { } empty) continue;
                    content = empty;
                }
                equations.Add(new DocxEquation(
                    content,
                    runs.Count, sourceRunIndex, display, alignment, equationRevision)
                {
                    ParentFontSize = ResolveRunProperties(null, paragraphStyleId, null, styles, tableCellStyle?.Run, out _).FontSize
                        ?? DocxDefaults.UnstyledRunFontSizePoints,
                    LimitSettings = (documentSettings ?? DocxDocumentSettings.Empty).MathLimitSettings,
                    DefaultAlignmentValue = documentSettings?.MathParagraphAlignmentValue,
                    HasRunRevisions = hasMathRunRevisions
                });
                inlineItems!.Add(new DocxInlineEquation(equations[^1]));
                RecordComplexFieldResult();
                if (firstTrailingBreak is { } trailing)
                {
                    foreach (var item in trailing.Run.ElementsAfterSelf().Prepend(trailing.Run))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        // Route range markers at their actual position in the
                        // paragraph stream. Only authored breaks advance lines.
                        bool mathRun = item.Name == MathNamespace + "r";
                        if (mathRun && !item.Elements().Any(IsDisplayMathBreak))
                            continue;
                        var inlineSource = mathRun ? AsParagraphRun(item, item == trailing.Run ? trailing.FirstBreakInRun : null) : item;
                        AddSourceContainerChild(inlineSource, source.Revision, removeDeletedText: false);
                    }
                }
            }

            DocxTextRun ReadMathRun(XElement run, string text)
            {
                DocxRevisionInfo? owner = equationRevision;
                XElement? wrapper = IsDisplayMathRun(run) ? null :
                    run.AncestorsAndSelf().TakeWhile(parent => parent.Name != MathNamespace + "oMath")
                    .FirstOrDefault(DocxSourceXml.IsRevisionContainer);
                if (wrapper is not null && !IsIncludedRevisionContainer(wrapper, markupMode)) owner = null;
                else if (wrapper is not null)
                {
                    mathRunRevisions ??= new();
                    if (!mathRunRevisions.TryGetValue(wrapper, out owner))
                    {
                        owner = CreateRevisionInfo(wrapper);
                        if (owner is not null) mathRunRevisions.Add(wrapper, owner);
                    }
                }
                if (owner is not null)
                {
                    hasMathRunRevisions = true;
                    AddRevision(paragraphRevisions, owner);
                }
                XElement? runProperties = DocxMathParser.ReadWordRunProperties(run);
                string? characterStyleId = ReadCharacterStyleIdFromProperties(runProperties);
                DocxResolvedRunProperties resolved = ResolveRunProperties(runProperties, paragraphStyleId,
                    characterStyleId, styles, tableCellStyle?.Run, out DocxRunStyleResolution resolution);
                return CreateResolvedTextRun(text, resolved, resolution,
                    complexScript: false, sourceRunIndex, sourceTextOffsetInRun: 0, owner, RevisionList(owner));
            }
        }

        bool AddSimpleField(XElement field, DocxRevisionInfo? revision)
        {
            string? instruction = (string?)field.Attribute(WordprocessingNamespace + "instr");
            DocxFieldKind kind = ResolveFieldKind(instruction);
            string? placeholder = ResolveSimpleFieldPlaceholder(field, kind);
            int fieldSourceRunIndex = sourceRunIndex;
            int fieldTextRunIndex = runs.Count;
            int fieldImageIndex = images.Count;
            int fieldBoxIndex = inlineTextBoxes.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            var fragment = field.Annotation<ComputedSimpleFieldFragment>();
            bool ownsCounter = fragment?.OwnsCounter != false;
            bool hasCachedResult = ownsCounter && FieldHasCachedResultContent(fragment?.Source ?? field);
            if (placeholder is not null)
            {
                // Bare computed values inherit their instruction's paragraph-style
                // formatting, rather than stale cached-result formatting. Explicit
                // MERGEFORMAT and unqualified switches keep the cached-style fallback.
                bool publishesCounter = ownsCounter && !HasReplacingFieldResult(complexFieldStack.Count);
                if (publishesCounter)
                {
                    XElement? formattingRun = null;
                    if (!UsesDefaultDynamicFieldFormatting(instruction, kind))
                    {
                        formattingRun = EnumerateRevisionScopedChildren((fragment?.Source ?? field).Elements(), markupMode,
                            cancellationToken, [WordprocessingNamespace + "r"], revision, inlineContent: true)
                            .FirstOrDefault()?.Element;
                    }
                    AddFieldPlaceholderRun(formattingRun ?? new XElement(WordprocessingNamespace + "r"), placeholder, fieldSourceRunIndex, revision, kind);
                }
                // The computed result replaces its cached runs and owns one
                // source position, including when there is no stored result.
                sourceRunIndex++;
                DocxTextRun? counterRun = publishesCounter ? runs[fieldTextRunIndex] : null;
                int? drawingSourceIndex = fieldSourceRunIndex;
                int resultTextOffset = 0, resultDrawingOrder = 0;
                int pendingRun = fieldTextRunIndex, pendingImage = fieldImageIndex, pendingBox = fieldBoxIndex;
                bool hasVmlStory = false;
                foreach (var storedScope in EnumerateRevisionScopedChildren(field.Elements(), markupMode, cancellationToken,
                    [WordprocessingNamespace + "r"], revision, inlineContent: true, includeNestedFieldResults: true))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    XElement resultRun = storedScope.Element;
                    var resultStyle = ResolveRunProperties(resultRun.Element(WordprocessingNamespace + "rPr"),
                        paragraphStyleId, ReadCharacterStyleId(resultRun), styles, tableCellStyle?.Run, out var resultResolution);
                    if (resultStyle.Hidden == true || SplitVmlTextBoxRun(resultRun) is not { } pieces)
                    {
                        AddCachedDrawings(resultRun, storedScope.Revision, resultStyle, resultResolution);
                        continue;
                    }

                    foreach (XElement piece in pieces)
                    {
                        if (piece.Name == WordprocessingNamespace + "txbxContent")
                        {
                            // A flattened box is placed between the cached drawings
                            // on each side. Publish the prefix before visiting it.
                            PublishInlineRunItems(pendingRun, pendingImage, pendingBox);
                            RetainDrawingRevision(storedScope.Revision);
                            AddVmlTextBoxContent(piece, storedScope.Revision);
                            pendingRun = runs.Count;
                            pendingImage = images.Count;
                            pendingBox = inlineTextBoxes.Count;
                            drawingSourceIndex = null;
                            resultTextOffset = 0;
                            resultDrawingOrder = 0;
                            hasVmlStory = true;
                        }
                        else
                        {
                            AddCachedDrawings(piece, storedScope.Revision, resultStyle, resultResolution);
                        }
                    }
                }
                if (ownsCounter)
                {
                    AddFieldReference(
                        kind,
                        DocxFieldSourceKind.Simple,
                        instruction,
                        placeholder,
                        fieldSourceRunIndex,
                        fieldTextRunIndex,
                        fieldTextLengthStart,
                        hasCachedResult,
                        rendersCachedResult: false,
                        usesPlaceholder: publishesCounter, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0,
                        fieldReferences: fieldReferences,
                        runs: runs,
                        counterReplacedByOuterField: !publishesCounter);
                }
                PublishInlineRunItems(pendingRun, pendingImage, pendingBox);
                return ownsCounter || hasVmlStory || images.Count != fieldImageIndex || inlineTextBoxes.Count != fieldBoxIndex;

                void AddCachedDrawings(XElement resultRun, DocxRevisionInfo? runRevision,
                    DocxResolvedRunProperties resultStyle, DocxRunStyleResolution resultResolution)
                {
                    int ownerIndex = drawingSourceIndex ?? sourceRunIndex;
                    RunDrawings stored = ReadRunDrawings(resultRun, ownerIndex, resultStyle.Hidden == true, runRevision);
                    if (stored.Pictures.Count != 0 || stored.TextBoxes.Count != 0)
                    {
                        RetainDrawingRevision(runRevision);
                        drawingSourceIndex ??= sourceRunIndex++;
                        DocxTextRun? pictureOwner = null;
                        if (stored.Pictures.Count != 0 && runRevision is { Kind: DocxRevisionKind.Insertion } owner)
                        {
                            pictureOwner = (owner == revision ? counterRun : null) ??
                                CreateResolvedTextRun(string.Empty, ApplyMarkupRevisionStyle(resultStyle, owner, markupMode),
                                    resultResolution, complexScript: false, ownerIndex, resultTextOffset, owner, RevisionList(owner));
                        }
                        var drawings = new List<DocxInlineItem>();
                        DocxInlineContent.AppendRun(drawings, [], stored.Pictures, cancellationToken, stored.TextBoxes);
                        foreach (var drawing in drawings)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (drawing is DocxInlinePicture picture)
                                images.Add(picture.Image with
                                {
                                    SourceTextOffsetInRun = resultTextOffset + picture.Image.SourceTextOffsetInRun,
                                    SourceOrderInRun = resultDrawingOrder++,
                                    SourceRunStyle = pictureOwner ?? picture.Image.SourceRunStyle
                                });
                            else if (drawing is DocxInlineBox box)
                                inlineTextBoxes.Add(box.Box with
                                {
                                    SourceTextOffsetInRun = resultTextOffset + box.Box.SourceTextOffsetInRun,
                                    SourceOrderInRun = resultDrawingOrder++
                                });
                        }
                    }
                    resultTextOffset += ReadRunText(resultRun).Length;
                }

                void RetainDrawingRevision(DocxRevisionInfo? drawingRevision)
                {
                    if (drawingRevision is not null && !ReferenceEquals(drawingRevision, revision) &&
                        !paragraphRevisions.Contains(drawingRevision, ReferenceEqualityComparer.Instance))
                        paragraphRevisions.Add(drawingRevision);
                }
            }

            bool hasSource = false;
            foreach (XElement fieldChild in field.Elements())
            {
                hasSource |= AddInlineContainerChild(fieldChild, revision);
            }

            AddFieldReference(
                kind,
                DocxFieldSourceKind.Simple,
                instruction,
                placeholder,
                fieldSourceRunIndex,
                fieldTextRunIndex,
                fieldTextLengthStart,
                hasCachedResult,
                rendersCachedResult: hasCachedResult, usesPlaceholder: false, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0,
                fieldReferences: fieldReferences,
                runs: runs);
            // Stored results own their parsed children. Empty/annotation-only
            // results keep field metadata without claiming neighboring ink.
            return hasSource;
        }

        DocxListLabel? ResolveParagraphListLabel()
        {
            var label = CreateListLabel(sourceNumbering, paragraphProperties, numbering, numberingCounters, styles, paragraphStyleId, tableCellStyle, paragraphMarkRun,
                DocxMarkupContext.FromMode(markupMode).AppliesInlineRevisionStyle,
                originalNumberingIncluded, numberingMarkInsertion);
            // A marker without its own paragraph/mark revision owns the resolved
            // paragraph color, not the first text run's authored or review ink.
            // Word inherits underline only from the numbering level. Keep the
            // separately qualified marker-revision path and other decorations.
            if (label is null || runs.Count == 0 || !knownNumberingMark || numberingMarkInsertion is not null ||
                (unmarkedListTextStyle is { } captured && captured.SourceRunIndex != runs[0].SourceRunIndex)) return label;
            return label with
            {
                DecorationSource = (unmarkedListTextStyle ?? runs[0]) with
                {
                    ColorHex = paragraphMarkRun.ColorHex,
                    Underline = false,
                    UnderlineValue = null,
                    UnderlineColorHex = null,
                    Revision = null,
                    Revisions = []
                }
            };
        }

        void CaptureUnmarkedListTextStyle(string text, DocxResolvedRunProperties properties,
            DocxRunStyleResolution resolution, int index, DocxRevisionInfo? revision)
        {
            // An enclosing paragraph revision also owns its marker. Only text-
            // scoped decoration is excluded from an otherwise unchanged label.
            if (runs.Count != 0 || text.Length == 0 || revision is null || paragraphMarkRevisions.Count != 0 ||
                ReferenceEquals(revision, inheritedRevision))
                return;
            bool complexScript = DocxScriptClassifier.IsComplexScriptRune(text.EnumerateRunes().First().Value);
            unmarkedListTextStyle = CreateResolvedTextRun(string.Empty, properties, resolution,
                complexScript, index, 0, null, null);
        }

        void AddFieldPlaceholderRun(XElement run, string text, int sourceRunIndex, DocxRevisionInfo? revision, DocxFieldKind? placeholderKind = null)
        {
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run,
                out DocxRunStyleResolution runStyleResolution);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadRunPropertyChangeRevisions(runProperties,
                paragraphStyleId, characterStyleId, styles, tableCellStyle?.Run, resolvedRun);
            AddRevisions(paragraphRevisions, runRevisions);
            CaptureUnmarkedListTextStyle(text, resolvedRun, runStyleResolution, sourceRunIndex, revision);
            AddResolvedTextRuns(runs, DocxTextRun.ResolveDisplayText(text, resolvedRun.AllCaps == true),
                ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode), runStyleResolution, sourceRunIndex,
                sourceTextOffsetInRun: 0, revision, MergeRevisionLists(revision, runRevisions), placeholderKind);
        }

        void AddParagraphRun(XElement run, DocxRevisionInfo? revision)
        {
            int firstRun = runs.Count;
            int firstReference = inlineReferences.Count;
            int firstImage = images.Count;
            int currentSourceRunIndex = sourceRunIndex;
            int firstBox = inlineTextBoxes.Count;
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties, paragraphStyleId, characterStyleId, styles, tableCellStyle?.Run,
                out DocxRunStyleResolution runStyleResolution);
            string text = ReadRunText(run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadRunPropertyChangeRevisions(runProperties,
                paragraphStyleId, characterStyleId, styles, tableCellStyle?.Run, resolvedRun);
            AddRevisions(paragraphRevisions, runRevisions);
            IReadOnlyList<DocxRevisionInfo> effectiveRevisions = MergeRevisionLists(revision, runRevisions);
            if (resolvedRun.Hidden != true && SplitVmlTextBoxRun(run) is { } pieces)
            {
                foreach (XElement piece in pieces)
                {
                    if (piece.Name == WordprocessingNamespace + "txbxContent")
                        AddVmlTextBoxContent(piece, revision);
                    else
                        AddParagraphRun(piece, revision);
                }
                if (sourceRunIndex == currentSourceRunIndex) sourceRunIndex++;
                RetainNonRenderingRunOwner(currentSourceRunIndex, runRevisions);
                return;
            }
            // Record authored picture children before processing field boundaries.
            bool hasComplexFieldMarkup = run.Elements().Any(IsComplexFieldMarkupElement);
            HashSet<XElement>? fieldPictureChildren = null;
            RunDrawings runDrawings = ReadRunDrawings(run, currentSourceRunIndex, resolvedRun.Hidden == true, revision,
                hasComplexFieldMarkup ? RecordPictureChild : null);
            if (runDrawings.TextBoxes.Count != 0) EnsureInlineContent();
            sourceRunIndex++;
            CaptureUnmarkedListTextStyle(text, resolvedRun, runStyleResolution, currentSourceRunIndex, revision);
            resolvedRun = ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode);
            // Content revisions need a painted or known nonprinting source.
            // Unknown objects cannot discharge an insertion through an empty run.
            // Run-formatting metadata is retained independently at publication.
            if (revision is not null && text.Length == 0 && (resolvedRun.Hidden == true || run.Elements().All(child =>
                child.Name == WordprocessingNamespace + "rPr" || IsComplexFieldMarkupElement(child) ||
                (child.Name == WordprocessingNamespace + "t" || child.Name == WordprocessingNamespace + "delText") && child.Value.Length == 0 ||
                tableCellStyle is not null && child.Name == WordprocessingNamespace + "br" && IsPageOrColumnBreak(child))))
            {
                RetainNonRenderingRunOwner(currentSourceRunIndex, effectiveRevisions);
            }

            if (hasComplexFieldMarkup)
            {
                runDrawings = AddOrderedRunContent(
                    run,
                    currentSourceRunIndex,
                    resolvedRun,
                    runStyleResolution,
                    revision,
                    effectiveRevisions,
                    runDrawings,
                    fieldPictureChildren);
                PublishRun(runDrawings);
                return;
            }

            // Nontext children delimit the text before and after their
            // source position, even when they share one formatted run.
            if (run.Elements().Any(child => IsInlineReferenceElement(child) || child.Annotation<AutomaticNoteStoryMarker>() is not null ||
                    child.Name == WordprocessingNamespace + "drawing" || child.Name == WordprocessingNamespace + "pict" ||
                    child.Name == WordprocessingNamespace + "sym"))
            {
                runDrawings = AddOrderedRunContent(run, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, effectiveRevisions,
                    runDrawings);
            }
            else if (text.Length != 0)
            {
                AddParagraphDisplayText(text, currentSourceRunIndex, sourceTextOffset: 0, resolvedRun, runStyleResolution, revision, effectiveRevisions);
                AddInlineReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, emitDisplayRuns: false, revision, effectiveRevisions);
            }
            else
            {
                AddInlineReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, emitDisplayRuns: false, revision, effectiveRevisions);
            }

            if (runDrawings.Pictures.Count != 0) RecordComplexFieldResult();
            PublishRun(runDrawings);

            void RecordPictureChild(XElement source)
            {
                XElement child = source.AncestorsAndSelf().First(element => element.Parent == run);
                (fieldPictureChildren ??= []).Add(child);
            }

            void PublishRun(RunDrawings drawings)
            {
                IReadOnlyList<DocxInlineImage> pictures = drawings.Pictures;
                inlineTextBoxes.AddRange(drawings.TextBoxes);
                if (inlineReferences.Count > firstReference &&
                    inlineReferences.Skip(firstReference).Any(reference => reference.CustomMarkText is not null))
                {
                    var owners = DocxCustomNoteMarkerOwners.IndexSourceRuns(runs.GetRange(firstRun, runs.Count - firstRun), cancellationToken);
                    foreach (var reference in inlineReferences.Skip(firstReference))
                        if (reference.SourceRunIndex == currentSourceRunIndex &&
                            DocxCustomNoteMarkerOwners.TryGetSourceRun(reference, owners, out var marker))
                        {
                            int index = runs.IndexOf(marker, firstRun);
                            if (index < 0) continue;
                            var marked = marker with { LiteralNoteMarkerOffset = reference.TextOffsetInRun };
                            runs[index] = marked;
                            owners[(marker.SourceRunIndex, marker.SourceTextOffsetInRun)] = marked;
                        }
                }
                if (pictures.Count != 0)
                {
                    if (revision?.Kind != DocxRevisionKind.Insertion) images.AddRange(pictures);
                    else
                    {
                        var style = runs.Count > firstRun ? runs[firstRun] : CreateResolvedTextRun(string.Empty,
                            resolvedRun, runStyleResolution, false, currentSourceRunIndex, 0, revision, effectiveRevisions);
                        foreach (var image in pictures)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            images.Add(image with { SourceRunStyle = style });
                        }
                    }
                }
                // Formatting belongs to the parsed run regardless of its payload.
                // Retain it after cached-result handling, which may already own
                // the complete revision set for this same source run.
                if (text.Length == 0) RetainNonRenderingRunOwner(currentSourceRunIndex, runRevisions);
                PublishInlineRunItems(firstRun, firstImage, firstBox);
            }
        }

        RunDrawings ReadRunDrawings(XElement run, int index, bool hidden, DocxRevisionInfo? revision,
            Action<XElement>? recordPictureSource = null)
        {
            if (!hidden) return ReadInlineDrawings(run, styles, numbering, documentSettings, package, relationships,
                markupMode, revision, cancellationToken, index, recordPictureSource);
            foreach (var inline in EnumerateRunOwnedDrawings(run, WordprocessingDrawingNamespace + "inline", cancellationToken))
                hiddenDrawingEquationSourceCount += InspectEquationSources(inline, markupMode, cancellationToken).Count;
            foreach (var textBox in EnumerateRunOwnedDrawings(run, VmlNamespace + "textbox", cancellationToken))
                hiddenDrawingEquationSourceCount += InspectEquationSources(textBox, markupMode, cancellationToken).Count;
            return new RunDrawings([], [], null);
        }

        IEnumerable<XElement>? SplitVmlTextBoxRun(XElement run)
        {
            if (run.Element(WordprocessingNamespace + "pict") is null) return null;
            var textBoxes = EnumerateRunOwnedDrawings(run, VmlNamespace + "textbox", cancellationToken)
                .Where(box => box.Ancestors(WordprocessingNamespace + "pict").Any(pict => ReferenceEquals(pict.Parent, run)))
                .ToLookup(box => box.Ancestors().First(owner => owner.Parent?.Name == WordprocessingNamespace + "pict" && ReferenceEquals(owner.Parent.Parent, run)));
            return textBoxes.Count == 0 ? null : Split();

            IEnumerable<XElement> Split()
            {
                // Split only at VML stories. Ordinary runs and the computed-field
                // drawing path keep their own decisions about the other children.
                var fragment = new List<XElement>();
                foreach (var child in run.Elements())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (child.Name == WordprocessingNamespace + "rPr") continue;
                    if (child.Name != WordprocessingNamespace + "pict") { fragment.Add(child); continue; }
                    foreach (var shape in child.Elements())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (!textBoxes.Contains(shape))
                        {
                            fragment.Add(new XElement(child.Name, child.Attributes(), new XElement(shape)));
                            continue;
                        }
                        if (fragment.Count != 0) yield return CopyFragment(fragment);
                        foreach (var box in textBoxes[shape])
                        foreach (var content in box.Elements(WordprocessingNamespace + "txbxContent"))
                            yield return content;
                    }
                }
                if (fragment.Count != 0) yield return CopyFragment(fragment);
            }

            XElement CopyFragment(List<XElement> fragment)
            {
                var view = new XElement(run.Name, run.Attributes(),
                    run.Element(WordprocessingNamespace + "rPr") is { } properties ? new XElement(properties) : null);
                foreach (var source in fragment)
                {
                    var copy = new XElement(source);
                    if (source.Annotation<AutomaticNoteStoryMarker>() is { } marker) copy.AddAnnotation(marker);
                    view.Add(copy);
                }
                fragment.Clear();
                return view;
            }
        }

        void AddVmlTextBoxContent(XElement content, DocxRevisionInfo? revision)
        {
            // The fallback shares placement with its anchor, but fields belong to the box story.
            var outerFields = complexFieldStack;
            complexFieldStack = [];
            try
            {
                bool first = true;
                foreach (var nested in EnumerateBookmarkedBlocks(content.Elements(), markupMode, cancellationToken, revision))
                {
                    if (nested.Element.Name != WordprocessingNamespace + "p") continue;
                    if (!first) AddParagraphRun(new XElement(WordprocessingNamespace + "r", new XElement(WordprocessingNamespace + "br")), revision);
                    first = false;
                    if (nested.LeadingBookmarks is { } leading)
                    {
                        // The paragraph now starts within the enclosing flow.
                        int textOffset = RelativeRunTextOffset(runs, sourceRunIndex);
                        int inlineIndex = GetInlineItemCount();
                        foreach (var anchor in leading)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            bookmarkAnchors.Add(anchor with { SourceRunIndex = sourceRunIndex,
                                TextRunIndex = runs.Count, TextOffset = textOffset, InlineItemIndex = inlineIndex });
                        }
                    }
                    var nestedParagraph = nested.Element;
                    var properties = nestedParagraph.Element(WordprocessingNamespace + "pPr");
                    if (nested.Revision is { } nestedRevision && nestedRevision != revision && !paragraphRevisions.Contains(nestedRevision))
                        paragraphRevisions.Add(nestedRevision);
                    AddRevisions(paragraphRevisions, ReadPropertyChangeRevisions(properties));
                    string? outerStyle = paragraphStyleId;
                    var outerCellStyle = tableCellStyle;
                    var outerListStyle = unmarkedListTextStyle;
                    try
                    {
                        // Textboxes own their run-style cascade, not the
                        // enclosing paragraph or table's formatting.
                        paragraphStyleId = ReadEffectiveParagraphStyleId(properties, styles);
                        tableCellStyle = null;
                        foreach (var item in nestedParagraph.Elements())
                            if (item.Name != WordprocessingNamespace + "pPr") AddInlineContainerChild(item, nested.Revision);
                    }
                    finally
                    {
                        paragraphStyleId = outerStyle;
                        tableCellStyle = outerCellStyle;
                        unmarkedListTextStyle = outerListStyle;
                    }
                }

                if (complexFieldStack.Count != 0)
                    FinalizeOpenComplexFields(continueInNextFragment: false);
            }
            finally
            {
                complexFieldStack = outerFields;
            }
        }

        void AddParagraphDisplayText(string text, int currentSourceRunIndex, int sourceTextOffset, DocxResolvedRunProperties resolvedRun, DocxRunStyleResolution runStyleResolution, DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo> effectiveRevisions)
        {
            if (text.Length == 0)
            {
                return;
            }

            string displayText = DocxTextRun.ResolveDisplayText(text, resolvedRun.AllCaps == true);
            DocxComplexFieldState[] resultFields = ActiveComplexResultFields();
            if (resultFields.Length != 0)
            {
                int activeFieldTextRunIndex = runs.Count;
                int activeFieldTextLengthStart = runs.Sum(run => run.Text.Length);
                foreach (DocxComplexFieldState field in resultFields)
                {
                    field.HasCachedResult = true;
                    field.EnsureTextSpan(activeFieldTextRunIndex, activeFieldTextLengthStart);
                }
            }

            // A computed result replaces its cached text, regardless of its
            // spelling or run boundaries. Pictures keep their separate source
            // ownership.
            if (resultFields.Any(field => field.PlaceholderEmitted))
            {
                // Replacing cache text does not discard its actual revision
                // owner or make unrelated paragraph content unsupported.
                RetainNonRenderingRunOwner(currentSourceRunIndex, effectiveRevisions);
                return;
            }

            int runsBefore = runs.Count;
            AddResolvedTextRuns(runs, displayText, resolvedRun, runStyleResolution, currentSourceRunIndex, sourceTextOffset, revision, effectiveRevisions);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
            }
        }

        void RetainNonRenderingRunOwner(int index, IReadOnlyList<DocxRevisionInfo> revisions)
        {
            if (revisions.Count == 0 || nonRenderingRunOwners is { Count: > 0 } &&
                nonRenderingRunOwners[^1].SourceRunIndex == index) return;
            (nonRenderingRunOwners ??= []).Add(new(index, revisions));
        }

        RunDrawings AddOrderedRunContent(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            RunDrawings drawings,
            IReadOnlySet<XElement>? fieldPictureChildren = null)
        {
            var positions = drawings.Positions;
            int[]? drawingTextOffsets = null;
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (fieldPictureChildren?.Contains(child) == true) RecordComplexFieldResult();
                DocxComplexFieldState? field = CurrentComplexField();
                bool isInstructionText = IsFieldInstructionText(child);
                // Outside a complex field's code region, instruction text follows the ordinary text rules.
                string childText = ReadRunTextChild(child, includeInstructionText: field is not { InResult: false });
                if (child.Name == WordprocessingNamespace + "fldChar")
                {
                    var position = field?.PendingPlaceholder is { SourceRunIndex: < 0 }
                        ? (Offset: textOffset, (positions ??= ReadRunInlinePositions(run, cancellationToken))[child].Order)
                        : (Offset: 0, Order: 0);
                    ApplyComplexFieldChar(child, currentSourceRunIndex, position, revision);
                }
                else if (isInstructionText && field is { InResult: false })
                {
                    if (IncludesFieldInstruction(child, field.UsesOriginalInstruction, revision?.Kind is
                        DocxRevisionKind.Insertion or DocxRevisionKind.MoveTo))
                    {
                        var position = (positions ??= ReadRunInlinePositions(run, cancellationToken))[child];
                        AddComplexFieldInstruction(field, child, currentSourceRunIndex, textOffset, position.Order,
                            resolvedRun, runStyleResolution, revision, revisions);
                    }
                }
                else if (IsInlineReferenceElement(child))
                {
                    var position = (positions ??= ReadRunInlinePositions(run, cancellationToken))[child];
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions,
                        inlineReferences: inlineReferences,
                        commentRanges: commentRanges,
                        openCommentRanges: openCommentRanges,
                        inlineReferenceCounters: inlineReferenceCounters,
                        documentSettings: documentSettings,
                        runs: runs, cancellationToken: cancellationToken, sourceDrawingIndex: position.Order);
                }
                else
                {
                    if (childText.Length != 0)
                    {
                        // The drawing scan knows source order before field parsing
                        // establishes which instruction elements carry literal text.
                        if (isInstructionText && drawings.Positions is { } drawingPositions)
                        {
                            drawingTextOffsets ??= new int[drawingPositions.Count];
                            drawingTextOffsets[drawingPositions[child].Order] += childText.Length;
                        }
                        var childProperties = resolvedRun;
                        if (child.Name == WordprocessingNamespace + "sym" && TryReadSymbol(child, out string? font, out _))
                        {
                            // A symbol selects a font-relative character, not a
                            // cased text spelling or an inherited font slot.
                            childProperties = resolvedRun with
                            {
                                FontFamily = font,
                                Fonts = new(font, font, font, font, null, null, null, null),
                                AllCaps = false,
                                SmallCaps = false
                            };
                        }
                        int markerRunsBefore = runs.Count;
                        AddParagraphDisplayText(
                            childText,
                            currentSourceRunIndex,
                            textOffset,
                            childProperties,
                            runStyleResolution,
                            revision,
                            revisions);
                        RetainAutomaticNoteStoryMarker(child, markerRunsBefore);
                    }
                }

                textOffset += childText.Length;
                childIndex++;
            }

            if (drawingTextOffsets is null) return drawings;

            for (int index = 1; index < drawingTextOffsets.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                drawingTextOffsets[index] += drawingTextOffsets[index - 1];
            }
            return new RunDrawings(
                drawings.Pictures.Select(picture => picture with
                {
                    SourceTextOffsetInRun = picture.SourceTextOffsetInRun + drawingTextOffsets[picture.SourceOrderInRun]
                }).ToArray(),
                drawings.TextBoxes.Select(box => box with
                {
                    SourceTextOffsetInRun = box.SourceTextOffsetInRun + drawingTextOffsets[box.SourceOrderInRun]
                }).ToArray(), drawings.Positions);
        }

        void ApplyComplexFieldChar(XElement fieldChar, int currentSourceRunIndex, (int Offset, int Order) position,
            DocxRevisionInfo? revision)
        {
            string? fieldCharType = (string?)fieldChar.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                bool unclosed = fieldChar.Annotation<UnclosedFieldBegin>() is not null ||
                    DocxSourceXml.FindFieldFlowOwner(fieldChar) is { } owner &&
                    ReadComplexFieldFlow(owner, markupMode, cancellationToken).Unclosed.Contains(fieldChar);
                complexFieldStack.Add(new DocxComplexFieldState(
                    currentSourceRunIndex,
                    runs.Count,
                    runs.Sum(run => run.Text.Length),
                    complexFieldStack.Count,
                    unclosed || OoxBoolean.IsTrue((string?)fieldChar.Attribute(WordprocessingNamespace + "fldLock")),
                    UsesOriginalFieldInstruction(markupMode, revision?.Kind == DocxRevisionKind.Deletion)));
                return;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is not null)
                {
                    EmitPendingFieldPlaceholder(field, currentSourceRunIndex, position.Offset, position.Order);
                    field.HasSeparate = true;
                    field.InResult = true;
                    field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
                }

                return;
            }

            if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is null)
                {
                    return;
                }

                EmitPendingFieldPlaceholder(field, currentSourceRunIndex, position.Offset, position.Order);
                complexFieldStack.RemoveAt(complexFieldStack.Count - 1);
                AddComplexFieldReference(field, null, null);
            }
        }

        void FinalizeOpenComplexFields(bool continueInNextFragment)
        {
            int firstRun = runs.Count;
            for (int fieldIndex = complexFieldStack.Count - 1; fieldIndex >= 0; fieldIndex--)
            {
                DocxComplexFieldState field = complexFieldStack[fieldIndex];
                cancellationToken.ThrowIfCancellationRequested();
                if (continueInNextFragment)
                {
                    // Preserve the unfinished instruction. An emitted value or
                    // stored result already has a placement owner in this frame.
                    if (field.PublishedReference is null &&
                        (field.PlaceholderEmitted || field.RendersCachedResult))
                    {
                        AddComplexFieldReference(field, null, null);
                        field.PublishedReference = (fieldReferences, fieldReferences.Count - 1);
                    }
                    field.ContinueInNextFragment();
                    continue;
                }

                EmitPendingFieldPlaceholder(field, Math.Max(0, sourceRunIndex - 1));
                string instruction = field.Instruction.ToString();
                string? placeholder = ResolveFieldPlaceholder(instruction, field.PreserveStoredResult);
                if (field.PublishedReference is not null || placeholder is not null ||
                    field.HasSeparate && field.HasCachedResult)
                {
                    AddComplexFieldReference(field, instruction, placeholder);
                }
            }

            if (!continueInNextFragment) complexFieldStack.Clear();
            PublishInlineRunItems(firstRun, images.Count);
        }

        void AddComplexFieldReference(DocxComplexFieldState field, string? instruction, string? placeholder)
        {
            instruction ??= field.Instruction.ToString();
            if (field.PublishedReference is { } owner)
            {
                // Update source-wide facts on the existing reference. Its run
                // span still indexes the fragment that actually painted it.
                owner.References[owner.Index] = owner.References[owner.Index] with
                {
                    Instruction = instruction,
                    HasSeparate = field.HasSeparate,
                    HasCachedResult = field.HasCachedResult,
                    RendersCachedResult = field.RendersCachedResult,
                    CounterReplacedByOuterField = field.CounterReplacedByOuterField,
                    InstructionRunCount = field.InstructionRunCount
                };
                return;
            }

            placeholder ??= ResolveFieldPlaceholder(instruction, field.PreserveStoredResult);
            int textRunIndex = field.ResultRunCount == 0 && !field.PlaceholderEmitted
                ? runs.Count
                : field.TextRunIndex;
            int textLengthStart = field.ResultRunCount == 0 && !field.PlaceholderEmitted
                ? runs.Sum(run => run.Text.Length)
                : field.TextLengthStart;
            AddFieldReference(
                ResolveFieldKind(instruction),
                DocxFieldSourceKind.ComplexInstruction,
                instruction,
                placeholder,
                field.InstructionSourceRunIndex >= 0 ? field.InstructionSourceRunIndex : field.SourceRunIndex,
                textRunIndex,
                textLengthStart,
                field.HasCachedResult,
                field.RendersCachedResult,
                field.PlaceholderEmitted,
                field.HasSeparate,
                field.NestingDepth,
                field.InstructionRunCount,
                field.ResultRunCount,
                fieldReferences: fieldReferences,
                runs: runs,
                counterReplacedByOuterField: field.CounterReplacedByOuterField);
        }

        void AddComplexFieldInstruction(
            DocxComplexFieldState field,
            XElement instruction,
            int currentSourceRunIndex,
            int sourceTextOffset,
            int sourceDrawingIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            string instructionText = (string?)instruction ?? string.Empty;
            if (field.InstructionSourceRunIndex < 0)
            {
                field.InstructionSourceRunIndex = currentSourceRunIndex;
            }

            field.Instruction.Append(instructionText);
            field.InstructionRunCount++;
            string? fieldPlaceholder = ResolveFieldPlaceholder(field.Instruction.ToString(), field.PreserveStoredResult);
            if (!field.PlaceholderEmitted)
            {
                // A run boundary does not finish the opcode: PAGE + REF is
                // PAGEREF. Keep the original formatting/source owner pending
                // until a delimiter or a field boundary confirms the name.
                if (fieldPlaceholder is null) field.PendingPlaceholder = null;
                else
                {
                    field.PendingPlaceholder ??= CreateResolvedTextRun(fieldPlaceholder, resolvedRun,
                        runStyleResolution, false, currentSourceRunIndex, sourceTextOffset, revision, revisions,
                        ResolveFieldKind(field.Instruction.ToString()))
                        with { SourceDrawingIndexInRun = sourceDrawingIndex };
                    if (ResolveFieldPlaceholder(field.Instruction.ToString(), field.PreserveStoredResult,
                        requireOpcodeTerminator: true) is not null)
                        EmitPendingFieldPlaceholder(field, currentSourceRunIndex, sourceTextOffset, sourceDrawingIndex);
                }
            }
        }

        void EmitPendingFieldPlaceholder(DocxComplexFieldState field, int currentSourceRunIndex, int offset = 0, int order = 0)
        {
            if (field.PendingPlaceholder is not { } pending) return;
            field.PendingPlaceholder = null;
            if (HasReplacingFieldResult(field.NestingDepth))
            {
                field.CounterReplacedByOuterField = true;
                return;
            }
            if (pending.SourceRunIndex < 0)
            {
                pending = pending with
                {
                    SourceRunIndex = currentSourceRunIndex,
                    SourceTextOffsetInRun = offset,
                    SourceDrawingIndexInRun = order
                };
                if (field.InstructionSourceRunIndex < 0) field.InstructionSourceRunIndex = currentSourceRunIndex;
            }
            field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
            runs.Add(pending);
            field.ResultRunCount++;
            field.PlaceholderEmitted = true;
        }

        // Computed results own their replacement text, including nested counters.
        // Stored fields and the separate drawing path retain their existing owners.
        bool HasReplacingFieldResult(int ancestorCount)
        {
            for (int index = ancestorCount - 1; index >= 0; index--)
            {
                if (complexFieldStack[index] is { InResult: true, PlaceholderEmitted: true })
                    return true;
            }
            return false;
        }

        void RecordComplexFieldResult()
        {
            foreach (var field in complexFieldStack)
            {
                if (!field.InResult) continue;
                field.HasCachedResult = true;
                // Computed counters keep their existing replacement owner.
                if (!field.PlaceholderEmitted) field.RendersCachedResult = true;
            }
        }

        DocxComplexFieldState? CurrentComplexField()
        {
            return complexFieldStack.Count == 0 ? null : complexFieldStack[^1];
        }

        DocxComplexFieldState[] ActiveComplexResultFields()
        {
            return complexFieldStack
                .Where(field => field.InResult)
                .ToArray();
        }

        void RetainAutomaticNoteStoryMarker(XElement child, int runsBefore)
        {
            if (child.Annotation<AutomaticNoteStoryMarker>() is null) return;
            for (int i = runsBefore; i < runs.Count; i++)
                runs[i] = runs[i] with { IsAutomaticNoteStoryMarker = true };
        }

        void AddInlineReferences(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            bool emitDisplayRuns,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRuns, revision, revisions,
                    inlineReferences: inlineReferences,
                    commentRanges: commentRanges,
                    openCommentRanges: openCommentRanges,
                    inlineReferenceCounters: inlineReferenceCounters,
                    documentSettings: documentSettings,
                    runs: runs, cancellationToken: cancellationToken);

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        static bool IsInlineReferenceElement(XElement element)
        {
            return ResolveInlineReferenceKind(element) is not null;
        }

    }

    private static DocxBookmarkAnchor ReadBookmarkAnchor(
        XElement bookmarkStart,
        int sourceRunIndex,
        int textRunIndex,
        int textOffset,
        int inlineItemIndex)
    {
        return new DocxBookmarkAnchor(
            (string?)bookmarkStart.Attribute(WordprocessingNamespace + "id"),
            (string?)bookmarkStart.Attribute(WordprocessingNamespace + "name"),
            sourceRunIndex,
            textRunIndex,
            textOffset) { InlineItemIndex = inlineItemIndex };
    }

    // RV06: comment and bookmark markers sit between runs while consumers resolve offsets
    // run-relatively, so paragraph-absolute capture sums are relativized here.
    private static int RelativeRunTextOffset(List<DocxTextRun> runs, int sourceRunIndex)
    {
        int absoluteOffset = runs.Sum(run => run.Text.Length);
        int prefixLength = 0;
        foreach (DocxTextRun run in runs)
        {
            if (run.SourceRunIndex < 0 || run.SourceRunIndex >= sourceRunIndex)
            {
                continue;
            }

            prefixLength += run.Text.Length;
        }

        return System.Math.Max(0, absoluteOffset - prefixLength);
    }

    private static void AddCommentRangeStart(
        XElement rangeStart,
        List<DocxCommentRangeStart> openCommentRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs,
        int inlineItemIndex)
    {
        openCommentRanges.Add(new DocxCommentRangeStart(
            OoxXml.NormalizeDecimalId((string?)rangeStart.Attribute(WordprocessingNamespace + "id")),
            sourceRunIndex,
            RelativeRunTextOffset(runs, sourceRunIndex))
        {
            InlineItemIndex = inlineItemIndex,
            FractionNumerator = rangeStart.Parent is { } parent && parent.Name == MathNamespace + "num" &&
                parent.Parent?.Name == MathNamespace + "f" ? parent : null
        });
    }

    private static void AddCommentRangeEnd(
        XElement rangeEnd,
        List<DocxCommentRangeStart> openCommentRanges,
        List<DocxCommentRange> commentRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs,
        int inlineItemIndex)
    {
        string? id = OoxXml.NormalizeDecimalId((string?)rangeEnd.Attribute(WordprocessingNamespace + "id"));
        int startIndex = openCommentRanges.FindLastIndex(start => string.Equals(start.Id, id, StringComparison.Ordinal));
        DocxCommentRangeStart? start = startIndex < 0 ? null : openCommentRanges[startIndex];
        if (startIndex >= 0)
        {
            openCommentRanges.RemoveAt(startIndex);
        }

        commentRanges.Add(new DocxCommentRange(
            id,
            start?.SourceRunIndex,
            start?.TextOffset,
            sourceRunIndex,
            RelativeRunTextOffset(runs, sourceRunIndex),
            ReferenceSourceRunIndex: null,
            ReferenceTextOffset: null)
        {
            StartInlineItemIndex = start?.InlineItemIndex,
            EndInlineItemIndex = inlineItemIndex,
            InFractionNumerator = start?.FractionNumerator is { } numerator && ReferenceEquals(numerator, rangeEnd.Parent)
        });
    }

    private static void AddRevisionMarker(
        XElement marker,
        List<DocxRevisionInfo> paragraphRevisions,
        List<DocxRevisionRangeStart> openRevisionRanges,
        List<DocxRevisionRange> revisionRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs,
        long contentSourcePosition,
        ref HashSet<DocxRevisionInfo>? nonPrintingRangeMarkers)
    {
        DocxRevisionInfo? markerRevision = CreateRevisionInfo(marker);
        AddRevision(paragraphRevisions, markerRevision);
        if (!DocxSourceXml.TryResolveRevisionRangeMarker(marker, out DocxRevisionKind? resolvedKind, out bool isStart) ||
            resolvedKind is not { } kind)
        {
            return;
        }

        if (isStart)
        {
            openRevisionRanges.Add(new DocxRevisionRangeStart(
                kind,
                (string?)marker.Attribute(WordprocessingNamespace + "id"),
                (string?)marker.Attribute(WordprocessingNamespace + "name"),
                (string?)marker.Attribute(WordprocessingNamespace + "author"),
                (string?)marker.Attribute(WordprocessingNamespace + "date"),
                sourceRunIndex,
                runs.Sum(run => run.Text.Length))
            {
                ContentSourcePosition = contentSourcePosition,
                IsNonPrintingBoundary = DocxSourceXml.IsNonPrintingRevisionRangeBoundary(marker),
                MarkerRevision = markerRevision
            });
            return;
        }

        string? id = (string?)marker.Attribute(WordprocessingNamespace + "id");
        string? decimalId = OoxXml.NormalizeDecimalId(id);
        int startIndex = openRevisionRanges.FindLastIndex(start =>
            start.Kind == kind &&
            string.Equals(OoxXml.NormalizeDecimalId(start.Id), decimalId, StringComparison.Ordinal));
        DocxRevisionRangeStart? startRange = startIndex < 0 ? null : openRevisionRanges[startIndex];
        if (startIndex >= 0)
        {
            openRevisionRanges.RemoveAt(startIndex);
        }

        int textOffset = runs.Sum(run => run.Text.Length);
        bool empty = startRange is { IsNonPrintingBoundary: true } && DocxSourceXml.IsNonPrintingRevisionRangeBoundary(marker) &&
            startRange.SourceRunIndex == sourceRunIndex && startRange.TextOffset == textOffset &&
            startRange.ContentSourcePosition == contentSourcePosition;
        revisionRanges.Add(new DocxRevisionRange(
            kind,
            id,
            startRange?.Name,
            startRange?.Author,
            startRange?.Date,
            startRange?.SourceRunIndex,
            startRange?.TextOffset,
            sourceRunIndex,
            textOffset) { HasContentSource = !empty });
        if (empty)
        {
            var markers = nonPrintingRangeMarkers ??= new(ReferenceEqualityComparer.Instance);
            if (startRange!.MarkerRevision is { } startRevision) markers.Add(startRevision);
            if (markerRevision is not null) markers.Add(markerRevision);
        }
    }

    private static void AddFieldReference(
        DocxFieldKind kind,
        DocxFieldSourceKind sourceKind,
        string? instruction,
        string? placeholder,
        int fieldSourceRunIndex,
        int fieldTextRunIndex,
        int fieldTextLengthStart,
        bool hasCachedResult,
        bool rendersCachedResult,
        bool usesPlaceholder,
        bool hasSeparate,
        int nestingDepth,
        int instructionRunCount,
        int resultRunCount,
        List<DocxFieldReference> fieldReferences,
        List<DocxTextRun> runs,
        bool counterReplacedByOuterField = false)
    {
        fieldReferences.Add(new DocxFieldReference(
            kind,
            sourceKind,
            instruction,
            placeholder,
            fieldSourceRunIndex,
            fieldTextRunIndex,
            runs.Count - fieldTextRunIndex,
            runs.Sum(run => run.Text.Length) - fieldTextLengthStart)
        {
            HasSeparate = hasSeparate,
            HasCachedResult = hasCachedResult,
            RendersCachedResult = rendersCachedResult,
            UsesPlaceholder = usesPlaceholder,
            CounterReplacedByOuterField = counterReplacedByOuterField,
            NestingDepth = nestingDepth,
            InstructionRunCount = instructionRunCount,
            ResultRunCount = resultRunCount
        });
    }

    private static DocxHyperlinkSpan AddHyperlinkSpan(
        XElement hyperlink,
        int sourceRunStartIndex,
        int sourceRunCount,
        int textRunStartIndex,
        int textRunCount,
        int inlineStart,
        int inlineEnd,
        List<DocxHyperlinkSpan> hyperlinkSpans,
        IReadOnlyDictionary<string, OoxRelationship> relationships, bool mathContent = false)
    {
        string? relationshipId = (string?)hyperlink.Attribute(RelationshipsNamespace + "id");
        relationships.TryGetValue(relationshipId ?? string.Empty, out OoxRelationship? relationship);
        var span = new DocxHyperlinkSpan(
            (string?)hyperlink.Attribute(WordprocessingNamespace + "anchor"),
            relationship?.Target,
            relationship?.TargetMode,
            relationship?.ResolvedTarget,
            sourceRunStartIndex,
            sourceRunCount,
            textRunStartIndex,
            textRunCount)
        {
            StartInlineItemIndex = inlineStart, EndInlineItemIndex = inlineEnd, IsMathContent = mathContent,
            IsMathRootContent = mathContent && hyperlink.Ancestors()
                .TakeWhile(ancestor => ancestor.Name != MathNamespace + "oMath")
                .All(ancestor => ancestor.Name.Namespace != MathNamespace)
        };
        hyperlinkSpans.Add(span);
        return span;
    }

    private static bool AddHyperlinkContainer(
        XElement hyperlink,
        DocxRevisionInfo? revision,
        Func<XElement, DocxRevisionInfo?, bool> addInlineChild,
        Func<int> getInlineItemCount,
        ref int sourceRunIndex,
        List<DocxTextRun> runs,
        List<DocxHyperlinkSpan> hyperlinkSpans,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        int sourceRunStartIndex = sourceRunIndex;
        int textRunStartIndex = runs.Count;
        int inlineStart = getInlineItemCount();
        bool hasSource = false;
        foreach (XElement hyperlinkChild in hyperlink.Elements())
        {
            hasSource |= addInlineChild(hyperlinkChild, revision);
        }

        AddHyperlinkSpan(
            hyperlink,
            sourceRunStartIndex,
            sourceRunIndex - sourceRunStartIndex,
            textRunStartIndex,
            runs.Count - textRunStartIndex,
            inlineStart,
            getInlineItemCount(),
            hyperlinkSpans,
            relationships);
        return hasSource;
    }

    private static DocxRelatedStoryKind? ResolveInlineReferenceKind(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "commentReference")
        {
            return DocxRelatedStoryKind.Comment;
        }

        if (element.Name == WordprocessingNamespace + "footnoteReference")
        {
            return DocxRelatedStoryKind.Footnote;
        }

        return element.Name == WordprocessingNamespace + "endnoteReference" ? DocxRelatedStoryKind.Endnote : null;
    }

    private static void AddInlineReference(
        XElement child,
        int currentSourceRunIndex,
        int childIndex,
        int textOffset,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution runStyleResolution,
        bool emitDisplayRun,
        DocxRevisionInfo? revision,
        IReadOnlyList<DocxRevisionInfo> revisions,
        List<DocxInlineReference> inlineReferences,
        List<DocxCommentRange> commentRanges,
        List<DocxCommentRangeStart> openCommentRanges,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        List<DocxTextRun> runs,
        CancellationToken cancellationToken,
        int sourceDrawingIndex = 0)
    {
        if (ResolveInlineReferenceKind(child) is not { } kind)
        {
            return;
        }

        string? customMarkFollows = kind == DocxRelatedStoryKind.Footnote || kind == DocxRelatedStoryKind.Endnote
            ? (string?)child.Attribute(WordprocessingNamespace + "customMarkFollows")
            : null;
        string? displayText = ResolveInlineReferenceDisplayText(kind, customMarkFollows);
        string? referenceId = OoxXml.NormalizeDecimalId((string?)child.Attribute(WordprocessingNamespace + "id"));
        inlineReferences.Add(new DocxInlineReference(
            kind,
            referenceId,
            customMarkFollows,
            displayText,
            currentSourceRunIndex,
            childIndex,
            textOffset)
        {
            CustomMarkText = ResolveCustomMarkText(),
            NoteNumberRestartValue = kind == DocxRelatedStoryKind.Footnote
                ? documentSettings?.FootnoteReferenceSettings.NumberRestartValue
                : kind == DocxRelatedStoryKind.Endnote ? documentSettings?.EndnoteReferenceSettings.NumberRestartValue : null,
            Revision = revision,
            Revisions = revisions
        });
        if (kind == DocxRelatedStoryKind.Comment)
        {
            AddCommentReferenceRange(referenceId, currentSourceRunIndex, textOffset);
        }

        if (emitDisplayRun && displayText is not null)
        {
            int firstRun = runs.Count;
            AddInlineReferenceDisplayRun(displayText, currentSourceRunIndex, textOffset, resolvedRun, runStyleResolution, revision);
            if (runs.Count == firstRun + 1)
            {
                runs[firstRun] = runs[firstRun] with
                { IsAutomaticNoteReferenceMarker = true, SourceDrawingIndexInRun = sourceDrawingIndex };
                inlineReferences[^1] = inlineReferences[^1] with { DisplayRun = runs[firstRun] };
            }
        }

        void AddCommentReferenceRange(string? id, int currentSourceRunIndex, int textOffset)
        {
            int rangeIndex = commentRanges.FindIndex(range =>
                string.Equals(range.Id, id, StringComparison.Ordinal) &&
                range.ReferenceSourceRunIndex is null);
            if (rangeIndex >= 0)
            {
                commentRanges[rangeIndex] = commentRanges[rangeIndex] with
                {
                    ReferenceSourceRunIndex = currentSourceRunIndex,
                    ReferenceTextOffset = textOffset
                };
                return;
            }

            int openRangeIndex = openCommentRanges.FindLastIndex(start => string.Equals(start.Id, id, StringComparison.Ordinal));
            if (openRangeIndex >= 0)
            {
                DocxCommentRangeStart start = openCommentRanges[openRangeIndex];
                openCommentRanges.RemoveAt(openRangeIndex);
                commentRanges.Add(new DocxCommentRange(
                    id,
                    start.SourceRunIndex,
                    start.TextOffset,
                    EndSourceRunIndex: null,
                    EndTextOffset: null,
                    currentSourceRunIndex,
                    textOffset) { StartInlineItemIndex = start.InlineItemIndex });
                return;
            }

            commentRanges.Add(new DocxCommentRange(
                id,
                StartSourceRunIndex: null,
                StartTextOffset: null,
                EndSourceRunIndex: null,
                EndTextOffset: null,
                currentSourceRunIndex,
                textOffset));
        }

        string? ResolveInlineReferenceDisplayText(DocxRelatedStoryKind kind, string? customMarkFollows)
        {
            if ((!string.IsNullOrEmpty(customMarkFollows) && !OoxBoolean.IsOff(customMarkFollows)) ||
                (kind != DocxRelatedStoryKind.Footnote && kind != DocxRelatedStoryKind.Endnote))
            {
                return null;
            }

            if (inlineReferenceCounters is null)
            {
                return null;
            }

            DocxNoteReferenceSettings settings = kind == DocxRelatedStoryKind.Endnote
                ? (documentSettings ?? DocxDocumentSettings.Empty).EndnoteReferenceSettings
                : (documentSettings ?? DocxDocumentSettings.Empty).FootnoteReferenceSettings;
            int next = inlineReferenceCounters.Next(kind, settings);
            // RV06 endnote probe: ECMA-376 defaults differ by kind (decimal footnotes,
            // lowerRoman endnotes); Office numbers the first style-less endnote "i".
            string? numberFormat = settings.NumberFormatValue
                ?? (kind == DocxRelatedStoryKind.Endnote ? "lowerRoman" : "decimal");
            return FormatNumberingCounter(next, numberFormat);
        }

        string? ResolveCustomMarkText()
        {
            // The mark owns the first resolved character in this source run.
            // Empty text and saved pagination hints contribute no character;
            // controls and other owners stop the lookup. Use emitted XML-space rules.
            if (!OoxBoolean.IsTrue(customMarkFollows)) return null;
            foreach (var following in child.ElementsAfterSelf())
            {
                cancellationToken.ThrowIfCancellationRequested();
                bool isText = following.Name == WordprocessingNamespace + "t";
                if (!isText && (following.HasElements || following.Value.Length != 0 ||
                    following.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration))) return null;
                string text = DocxTextRun.ResolveDisplayText(ReadRunTextChild(following), resolvedRun.AllCaps == true);
                if (text.Length != 0)
                    return char.IsWhiteSpace(text[0]) || char.IsSurrogate(text[0]) ? null : text[..1];
                if (following.HasElements || !isText && following.Name != WordprocessingNamespace + "lastRenderedPageBreak") return null;
            }
            return null;
        }

        void AddInlineReferenceDisplayRun(
            string displayText,
            int currentSourceRunIndex,
            int textOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision)
        {
            // RV06 footnote-align probe: Office renders the mark at baseline when no
            // superscript is specified (style-less reference exports full-size marks
            // on the baseline), so the display run follows the resolved cascade
            // instead of forcing superscript. Documents whose FootnoteReference or
            // EndnoteReference style (or direct run properties) specify superscript
            // keep it through the cascade.
            AddResolvedTextRuns(
                runs,
                DocxTextRun.ResolveDisplayText(displayText, resolvedRun.AllCaps == true),
                resolvedRun,
                runStyleResolution,
                currentSourceRunIndex,
                textOffset,
                revision,
                revisions: null);
        }
    }

    private static void AddResolvedTextRuns(List<DocxTextRun> runs, string text, DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution, int sourceRunIndex, int sourceTextOffsetInRun,
        DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo>? revisions, DocxFieldKind? fieldKind = null)
    {
        var segment = new StringBuilder();
        DocxFontSlot? currentSlot = null;
        int segmentSourceOffset = sourceTextOffsetInRun;
        foreach (Rune rune in text.EnumerateRunes())
        {
            DocxFontSlot slot = DocxScriptClassifier.FontSlot(rune.Value);
            if (currentSlot is not null && currentSlot.Value != slot &&
                !DocxFontResolver.SameTypefaceSelection(resolvedRun.Fonts, resolvedRun.FontFamily, currentSlot.Value, slot))
            {
                string segmentText = segment.ToString();
                AddResolvedTextRun(runs, segmentText, resolvedRun, styleResolution,
                    currentSlot.Value == DocxFontSlot.ComplexScript, sourceRunIndex, segmentSourceOffset,
                    revision, revisions, fieldKind);
                segmentSourceOffset += segmentText.Length;
                segment.Clear();
            }

            segment.Append(rune);
            currentSlot = slot;
        }

        if (segment.Length != 0 && currentSlot is not null)
        {
            AddResolvedTextRun(runs, segment.ToString(), resolvedRun, styleResolution,
                currentSlot.Value == DocxFontSlot.ComplexScript, sourceRunIndex, segmentSourceOffset,
                revision, revisions, fieldKind);
        }
    }

    private static void AddResolvedTextRun(List<DocxTextRun> runs, string text, DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution, bool complexScript, int sourceRunIndex, int sourceTextOffsetInRun,
        DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo>? revisions, DocxFieldKind? fieldKind = null)
    {
        runs.Add(CreateResolvedTextRun(text, resolvedRun, styleResolution, complexScript, sourceRunIndex,
            sourceTextOffsetInRun, revision, revisions, fieldKind));
    }

    private static DocxTextRun CreateResolvedTextRun(string text, DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution, bool complexScript, int sourceRunIndex, int sourceTextOffsetInRun,
        DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo>? revisions, DocxFieldKind? fieldKind = null)
    {
        bool bold = complexScript
            ? resolvedRun.ComplexScriptBold ?? resolvedRun.Bold ?? false
            : resolvedRun.Bold ?? false;
        bool italic = complexScript
            ? resolvedRun.ComplexScriptItalic ?? resolvedRun.Italic ?? false
            : resolvedRun.Italic ?? false;
        string? fontFamily = complexScript
            ? FirstNonEmpty(resolvedRun.Fonts.ComplexScript, resolvedRun.FontFamily)
            : resolvedRun.FontFamily;
        bool isScriptShiftedRun =
            DocxTextRun.ParseVerticalAlignment(resolvedRun.VerticalAlignmentValue) is not DocxRunVerticalAlignment.Baseline;
        return new DocxTextRun(
            text,
            resolvedRun.FontSize ?? DocxDefaults.UnstyledRunFontSizePoints,
            resolvedRun.ColorHex,
            bold,
            italic,
            resolvedRun.Underline ?? false,
            resolvedRun.UnderlineValue,
            fontFamily,
            resolvedRun.CharacterSpacingPoints ?? 0d,
            resolvedRun.AllCaps ?? false,
            resolvedRun.VerticalAlignmentValue,
            resolvedRun.Strike ?? false,
            resolvedRun.StrikeValue,
            resolvedRun.DoubleStrike ?? false,
            resolvedRun.DoubleStrikeValue,
            resolvedRun.HighlightValue,
            resolvedRun.ShadingFillHex,
            resolvedRun.ShadingValue,
            resolvedRun.ShadingColor,
            resolvedRun.SmallCaps ?? false,
            resolvedRun.SmallCapsValue,
            resolvedRun.Hidden ?? false,
            resolvedRun.HiddenValue,
            resolvedRun.UnderlineColorHex)
        {
            Fonts = resolvedRun.Fonts,
            KerningMinimumHalfPointsValue = resolvedRun.KerningMinimumHalfPointsValue,
            FieldKind = fieldKind,
            ScriptBaseFontSize = isScriptShiftedRun && resolvedRun.FontSize is null ? DocxDefaults.FontSizePoints : null,
            StyleResolution = styleResolution,
            SourceRunIndex = sourceRunIndex,
            SourceTextOffsetInRun = sourceTextOffsetInRun,
            Revision = revision,
            Revisions = revisions ?? RevisionList(revision)
        };
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    private static string ReadRunText(XElement run)
    {
        var text = new System.Text.StringBuilder();
        foreach (XElement child in run.Elements())
        {
            text.Append(ReadRunTextChild(child));
        }

        return text.ToString();
    }

    internal static bool TryReadSymbol(XElement symbol, [NotNullWhen(true)] out string? font, out char character)
    {
        font = (string?)symbol.Attribute(WordprocessingNamespace + "font");
        ReadOnlySpan<char> code = ((string?)symbol.Attribute(WordprocessingNamespace + "char"))
            .AsSpan().Trim(" \t\r\n".AsSpan());
        if (!string.IsNullOrWhiteSpace(font) && code.Length == 4 &&
            ushort.TryParse(code, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out ushort value) &&
            Rune.TryCreate(value, out Rune rune) &&
            Rune.GetUnicodeCategory(rune) is not (UnicodeCategory.Control or UnicodeCategory.Format or
                UnicodeCategory.LineSeparator or UnicodeCategory.ParagraphSeparator))
        {
            character = (char)value;
            return true;
        }
        // An unsupported payload still owns one visible replacement/source slot.
        // The shared part-level probe diagnoses it; ordinary font recovery owns
        // unavailable fonts and glyphs for otherwise well-formed symbols.
        font = null;
        character = '\uFFFD';
        return false;
    }

    private static string ReadRunTextChild(XElement child, bool includeInstructionText = false)
    {
        if (child.Name == WordprocessingNamespace + "sym")
        {
            TryReadSymbol(child, out _, out char character);
            return character.ToString();
        }
        if (child.Name == WordprocessingNamespace + "t" ||
            child.Name == WordprocessingNamespace + "delText" ||
            includeInstructionText && IsFieldInstructionText(child))
        {
            string value = child.Value;
            // Word 16 honors preservation on the text element itself, not on
            // ancestor runs/paragraphs. Unmarked boundary XML whitespace is
            // insignificant; interior spaces and nonbreaking spaces remain.
            if ((string?)child.Attribute(XNamespace.Xml + "space") != "preserve")
            {
                int start = 0;
                int end = value.Length;
                while (start < end && IsXmlWhitespace(value[start]))
                {
                    start++;
                }
                while (end > start && IsXmlWhitespace(value[end - 1]))
                {
                    end--;
                }
                if (start != 0 || end != value.Length)
                {
                    value = value.Substring(start, end - start);
                }
            }
            // Literal controls in w:t are text whitespace. w:tab and w:br
            // below carry the actual tab-stop and line-break semantics.
            return value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
        }

        if (child.Name == WordprocessingNamespace + "tab")
        {
            return "\t";
        }

        if (child.Name == WordprocessingNamespace + "noBreakHyphen")
        {
            return "\u2011";
        }

        if (child.Name == WordprocessingNamespace + "softHyphen")
        {
            return "\u00AD";
        }

        if (DocxSourceXml.IsTextWrappingBreak(child))
        {
            return "\n";
        }

        return string.Empty;
    }

    private static bool IsXmlWhitespace(char value) => value is ' ' or '\t' or '\r' or '\n';

    private static string? ReadParagraphStyleId(XElement? paragraphProperties)
    {
        return (string?)paragraphProperties?
            .Element(WordprocessingNamespace + "pStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    private static string? ReadEffectiveParagraphStyleId(XElement? properties, DocxStyleSet styles) =>
        properties?.Element(WordprocessingNamespace + "pStyle") is null
            ? styles.DeclaredDefaultParagraphStyleId : ReadParagraphStyleId(properties);
}

