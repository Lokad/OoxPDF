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
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle? tableCellStyle,
        Dictionary<DocxRelatedStoryKind, int>? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        DocxRevisionInfo? inheritedRevision,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        string? paragraphStyleId = ReadParagraphStyleId(paragraphProperties);
        DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
            paragraphProperties,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        DocxParagraphStyleResolution styleResolution = CreateParagraphStyleResolution(
            paragraphProperties,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        var runs = new List<DocxTextRun>();
        var images = new List<DocxInlineImage>();
        var inlineReferences = new List<DocxInlineReference>();
        var commentRanges = new List<DocxCommentRange>();
        var openCommentRanges = new List<DocxCommentRangeStart>();
        var revisionRanges = new List<DocxRevisionRange>();
        var openRevisionRanges = new List<DocxRevisionRangeStart>();
        var fieldReferences = new List<DocxFieldReference>();
        var hyperlinkSpans = new List<DocxHyperlinkSpan>();
        var bookmarkAnchors = new List<DocxBookmarkAnchor>();
        var paragraphRevisions = new List<DocxRevisionInfo>();
        AddRevision(paragraphRevisions, inheritedRevision);
        AddRevisions(paragraphRevisions, ReadPropertyChangeRevisions(paragraphProperties));
        XElement? paragraphMarkRunProperties = paragraphProperties?.Element(WordprocessingNamespace + "rPr");
        IReadOnlyList<DocxRevisionInfo> paragraphMarkRevisions = ReadPropertyChangeRevisions(paragraphMarkRunProperties);
        bool hasDeletedParagraphMark = paragraphMarkRevisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion && revision.SourceElement == "del");
        AddRevisions(paragraphRevisions, paragraphMarkRevisions);
        bool pageInstructionSeen = false;
        var complexFieldStack = new List<DocxComplexFieldState>();
        int sourceRunIndex = 0;
        foreach (XElement child in paragraph.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child, inheritedRevision);
            }
            else if (IsRevisionContainer(child))
            {
                AddRevisionRunContainer(child, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                AddHyperlinkContainer(child, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "sdt")
            {
                AddContentControl(child, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child, bookmarkAnchors, sourceRunIndex, runs);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeStart")
            {
                AddCommentRangeStart(child, openCommentRanges, sourceRunIndex, runs);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                AddCommentRangeEnd(child, openCommentRanges, commentRanges, sourceRunIndex, runs);
            }
            else if (IsRevisionMarkerElement(child))
            {
                AddRevisionMarker(child);
            }
        }

        void AddInlineContainerChild(XElement child, DocxRevisionInfo? revision)
        {
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen, revision);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child, bookmarkAnchors, sourceRunIndex, runs);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child, revision);
            }
            else if (IsRevisionContainer(child))
            {
                AddRevisionRunContainer(child, inheritedRevision: revision);
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                AddHyperlinkContainer(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "sdt")
            {
                AddContentControl(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "sdtContent")
            {
                AddVisibleRunContainer(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeStart")
            {
                AddCommentRangeStart(child, openCommentRanges, sourceRunIndex, runs);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                AddCommentRangeEnd(child, openCommentRanges, commentRanges, sourceRunIndex, runs);
            }
            else if (IsRevisionMarkerElement(child))
            {
                AddRevisionMarker(child);
            }
        }

        void AddVisibleRunContainer(XElement container, DocxRevisionInfo? revision)
        {
            foreach (XElement containerChild in container.Elements())
            {
                AddInlineContainerChild(containerChild, revision);
            }
        }

        void AddContentControl(XElement contentControl, DocxRevisionInfo? revision)
        {
            foreach (XElement content in contentControl.Elements(WordprocessingNamespace + "sdtContent"))
            {
                AddVisibleRunContainer(content, revision);
            }
        }

        void AddRevisionRunContainer(XElement container, DocxRevisionInfo? inheritedRevision)
        {
            if (!IsIncludedRevisionContainer(container, markupMode))
            {
                return;
            }

            DocxRevisionInfo? revision = CreateRevisionInfo(container) ?? inheritedRevision;
            if (revision is not null)
            {
                paragraphRevisions.Add(revision);
            }

            AddVisibleRunContainer(container, revision);
        }

        FinalizeOpenComplexFields();

        if (runs.Count == 0 && images.Count == 0)
        {
            DocxResolvedRunProperties paragraphMarkRun = ResolveRunProperties(
                paragraphMarkRunProperties,
                paragraphStyleId,
                characterStyleId: null,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution paragraphMarkStyleResolution = CreateRunStyleResolution(
                paragraphMarkRunProperties,
                paragraphStyleId,
                characterStyleId: null,
                styles,
                tableCellStyle?.Run);
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
        double lineSpacingFactor = resolvedParagraph.LineSpacingFactor ?? ResolveDefaultAutoLineSpacingFactor(resolvedParagraph);
        double paragraphLineHeight = resolvedParagraph.LineSpacingPoints ?? paragraphFontSize * lineSpacingFactor;

        return new DocxParagraph(
            runs,
            images,
            paragraphStyleId,
            resolvedParagraph.Alignment ?? DocxTextAlignment.Left,
            resolvedParagraph.AlignmentValue,
            ResolveSpacingBeforePoints(resolvedParagraph, paragraphLineHeight),
            ResolveSpacingAfterPoints(resolvedParagraph, paragraphLineHeight),
            lineSpacingFactor,
            resolvedParagraph.LineSpacingPoints,
            resolvedParagraph.Spacing,
            resolvedParagraph.KeepRules,
            CreateListLabel(paragraphProperties, numbering, numberingCounters))
        {
            Indent = resolvedParagraph.Indent,
            TabStops = resolvedParagraph.TabStops,
            SnapToGrid = resolvedParagraph.SnapToGrid,
            SnapToGridValue = resolvedParagraph.SnapToGridValue,
            WordWrap = resolvedParagraph.WordWrap,
            WordWrapValue = resolvedParagraph.WordWrapValue,
            StyleResolution = styleResolution,
            InlineReferences = inlineReferences,
            CommentRanges = commentRanges,
            RevisionRanges = revisionRanges,
            FieldReferences = fieldReferences,
            Hyperlinks = hyperlinkSpans,
            BookmarkAnchors = bookmarkAnchors,
            Revisions = paragraphRevisions,
            HasDeletedParagraphMark = hasDeletedParagraphMark
        };

        void AddRevisionMarker(XElement marker)
        {
            AddRevision(paragraphRevisions, CreateRevisionInfo(marker));
            if (!TryResolveRevisionRangeMarker(marker, out DocxRevisionKind? resolvedKind, out bool isStart) ||
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
                    runs.Sum(run => run.Text.Length)));
                return;
            }

            string? id = (string?)marker.Attribute(WordprocessingNamespace + "id");
            int startIndex = openRevisionRanges.FindLastIndex(start =>
                start.Kind == kind &&
                string.Equals(start.Id, id, StringComparison.Ordinal));
            DocxRevisionRangeStart? startRange = startIndex < 0 ? null : openRevisionRanges[startIndex];
            if (startIndex >= 0)
            {
                openRevisionRanges.RemoveAt(startIndex);
            }

            revisionRanges.Add(new DocxRevisionRange(
                kind,
                id,
                startRange?.Name,
                startRange?.Author,
                startRange?.Date,
                startRange?.SourceRunIndex,
                startRange?.TextOffset,
                sourceRunIndex,
                runs.Sum(run => run.Text.Length)));
        }

        void AddHyperlinkSpan(
            XElement hyperlink,
            int sourceRunStartIndex,
            int sourceRunCount,
            int textRunStartIndex,
            int textRunCount,
            int textLength)
        {
            string? relationshipId = (string?)hyperlink.Attribute(RelationshipsNamespace + "id");
            relationships.TryGetValue(relationshipId ?? string.Empty, out OoxRelationship? relationship);
            hyperlinkSpans.Add(new DocxHyperlinkSpan(
                relationshipId,
                (string?)hyperlink.Attribute(WordprocessingNamespace + "anchor"),
                (string?)hyperlink.Attribute(WordprocessingNamespace + "tooltip"),
                (string?)hyperlink.Attribute(WordprocessingNamespace + "history"),
                relationship?.Target,
                relationship?.TargetMode,
                relationship?.ResolvedTarget,
                sourceRunStartIndex,
                sourceRunCount,
                textRunStartIndex,
                textRunCount,
                textLength));
        }

        void AddSimpleField(XElement field, DocxRevisionInfo? revision)
        {
            string? instruction = (string?)field.Attribute(WordprocessingNamespace + "instr");
            DocxFieldKind kind = ResolveFieldKind(instruction);
            string? placeholder = ResolveFieldPlaceholder(instruction);
            int fieldSourceRunIndex = sourceRunIndex;
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            bool hasCachedResult = FieldHasCachedResultText(field);
            if (placeholder is not null)
            {
                XElement? firstRun = field.Elements(WordprocessingNamespace + "r").FirstOrDefault();
                if (firstRun is null)
                {
                    runs.Add(new DocxTextRun(placeholder, DocxDefaults.FontSizePoints, null, false, false, false, null, null)
                    {
                        SourceRunIndex = fieldSourceRunIndex,
                        Revision = revision,
                        Revisions = RevisionList(revision)
                    });
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
                        usesPlaceholder: true, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0);
                    return;
                }

                AddFieldPlaceholderRun(firstRun, placeholder, fieldSourceRunIndex, revision);
                images.AddRange(ReadInlineImages(firstRun, package, relationships, revision));
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
                    usesPlaceholder: true, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0);
                return;
            }

            foreach (XElement fieldChild in field.Elements())
            {
                AddInlineContainerChild(fieldChild, revision);
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
                rendersCachedResult: hasCachedResult, usesPlaceholder: false, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0);
        }

        void AddFieldReference(
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
            int resultRunCount)
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
                NestingDepth = nestingDepth,
                InstructionRunCount = instructionRunCount,
                ResultRunCount = resultRunCount
            });
        }

        void AddHyperlinkContainer(XElement hyperlink, DocxRevisionInfo? revision)
        {
            int sourceRunStartIndex = sourceRunIndex;
            int textRunStartIndex = runs.Count;
            int textLengthStart = runs.Sum(run => run.Text.Length);
            foreach (XElement hyperlinkChild in hyperlink.Elements())
            {
                AddInlineContainerChild(hyperlinkChild, revision);
            }

            AddHyperlinkSpan(
                hyperlink,
                sourceRunStartIndex,
                sourceRunIndex - sourceRunStartIndex,
                textRunStartIndex,
                runs.Count - textRunStartIndex,
                runs.Sum(run => run.Text.Length) - textLengthStart);
        }

        void AddFieldPlaceholderRun(XElement run, string text, int sourceRunIndex, DocxRevisionInfo? revision)
        {
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution runStyleResolution = CreateRunStyleResolution(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadPropertyChangeRevisions(runProperties);
            AddRevisions(paragraphRevisions, runRevisions);
            AddResolvedTextRuns(
                runs,
                resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text,
                ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode),
                runStyleResolution,
                sourceRunIndex,
                sourceTextOffsetInRun: 0,
                revision,
                MergeRevisionLists(revision, runRevisions));
        }

        void AddParagraphRun(XElement run, ref bool currentPageInstructionSeen, DocxRevisionInfo? revision)
        {
            int currentSourceRunIndex = sourceRunIndex++;
            string text = ReadRunText(run);
            string? fieldInstruction = run
                .Elements(WordprocessingNamespace + "instrText")
                .Select(instruction => (string?)instruction)
                .FirstOrDefault(value => value is not null);
            string? placeholder = ResolveFieldPlaceholder(fieldInstruction);
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution runStyleResolution = CreateRunStyleResolution(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadPropertyChangeRevisions(runProperties);
            AddRevisions(paragraphRevisions, runRevisions);
            IReadOnlyList<DocxRevisionInfo> effectiveRevisions = MergeRevisionLists(revision, runRevisions);
            resolvedRun = ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode);

            if (run.Elements().Any(IsComplexFieldMarkupElement))
            {
                AddComplexFieldAwareRun(
                    run,
                    currentSourceRunIndex,
                    resolvedRun,
                    runStyleResolution,
                    revision,
                    effectiveRevisions,
                    ref currentPageInstructionSeen);
                images.AddRange(ReadInlineImages(run, package, relationships, revision));
                return;
            }

            if (placeholder is not null)
            {
                text = placeholder;
                currentPageInstructionSeen = true;
            }
            else if (currentPageInstructionSeen && text.Trim().All(char.IsDigit))
            {
                foreach (DocxComplexFieldState field in ActiveComplexResultFields())
                {
                    field.HasCachedResult = true;
                }

                text = string.Empty;
            }

            if (placeholder is null &&
                fieldInstruction is null &&
                run.Elements().Any(IsInlineReferenceElement))
            {
                AddOrderedRunTextAndReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, effectiveRevisions);
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

            foreach (DocxVmlTextBoxContent textBoxContent in ReadVmlTextBoxContents(run))
            {
                AddRevisions(paragraphRevisions, textBoxContent.Revisions);
                AddParagraphDisplayText(textBoxContent.Text, currentSourceRunIndex, sourceTextOffset: 0, resolvedRun, runStyleResolution, revision, effectiveRevisions);
            }

            if (fieldInstruction is not null)
            {
                AddFieldReference(
                    ResolveFieldKind(fieldInstruction),
                    DocxFieldSourceKind.ComplexInstruction,
                    fieldInstruction,
                    placeholder,
                    currentSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart, hasCachedResult: false, rendersCachedResult: false, usesPlaceholder: false, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0);
            }

            images.AddRange(ReadInlineImages(run, package, relationships, revision));
        }

        void AddParagraphDisplayText(
            string text,
            int currentSourceRunIndex,
            int sourceTextOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> effectiveRevisions)
        {
            if (text.Length == 0)
            {
                return;
            }

            string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
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

            int runsBefore = runs.Count;
            AddResolvedTextRuns(runs, displayText, resolvedRun, runStyleResolution, currentSourceRunIndex, sourceTextOffset, revision, effectiveRevisions);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
            }
        }

        void AddComplexFieldAwareRun(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (child.Name == WordprocessingNamespace + "fldChar")
                {
                    ApplyComplexFieldChar(child, currentSourceRunIndex);
                }
                else if (child.Name == WordprocessingNamespace + "instrText")
                {
                    AddComplexFieldInstruction(child, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, revisions, ref currentPageInstructionSeen);
                }
                else if (IsInlineReferenceElement(child))
                {
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions);
                }
                else
                {
                    string childText = ReadRunTextChild(child);
                    if (childText.Length != 0)
                    {
                        AddComplexFieldText(
                            childText,
                            currentSourceRunIndex,
                            textOffset,
                            resolvedRun,
                            runStyleResolution,
                            revision,
                            revisions,
                            ref currentPageInstructionSeen);
                    }
                }

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        void ApplyComplexFieldChar(XElement fieldChar, int currentSourceRunIndex)
        {
            string? fieldCharType = (string?)fieldChar.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                complexFieldStack.Add(new DocxComplexFieldState(
                    currentSourceRunIndex,
                    runs.Count,
                    runs.Sum(run => run.Text.Length),
                    complexFieldStack.Count));
                return;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is not null)
                {
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

                complexFieldStack.RemoveAt(complexFieldStack.Count - 1);
                AddComplexFieldReference(field, null, null);
            }
        }

        void FinalizeOpenComplexFields()
        {
            for (int fieldIndex = complexFieldStack.Count - 1; fieldIndex >= 0; fieldIndex--)
            {
                DocxComplexFieldState field = complexFieldStack[fieldIndex];
                string instruction = field.Instruction.ToString();
                string? placeholder = ResolveFieldPlaceholder(instruction);
                if (placeholder is null && (!field.HasSeparate || !field.HasCachedResult))
                {
                    continue;
                }

                AddComplexFieldReference(field, instruction, placeholder);
            }

            complexFieldStack.Clear();
        }

        void AddComplexFieldReference(DocxComplexFieldState field, string? instruction, string? placeholder)
        {
            instruction ??= field.Instruction.ToString();
            placeholder ??= ResolveFieldPlaceholder(instruction);
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
                field.ResultRunCount);
        }

        void AddComplexFieldInstruction(
            XElement instruction,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            string instructionText = (string?)instruction ?? string.Empty;
            DocxComplexFieldState? field = CurrentComplexField();
            if (field is null)
            {
                int fieldTextRunIndex = runs.Count;
                int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                string? placeholder = ResolveFieldPlaceholder(instructionText);
                if (placeholder is not null)
                {
                    AddResolvedTextRuns(
                        runs,
                        resolvedRun.AllCaps == true ? placeholder.ToUpperInvariant() : placeholder,
                        resolvedRun,
                        runStyleResolution,
                        currentSourceRunIndex,
                        sourceTextOffsetInRun: 0,
                        revision,
                        revisions);
                    currentPageInstructionSeen = true;
                }

                AddFieldReference(
                    ResolveFieldKind(instructionText),
                    DocxFieldSourceKind.ComplexInstruction,
                    instructionText,
                    placeholder,
                    currentSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart,
                    usesPlaceholder: placeholder is not null,
                    instructionRunCount: 1, hasCachedResult: false, rendersCachedResult: false, hasSeparate: false, nestingDepth: 0, resultRunCount: 0);
                return;
            }

            if (field.InstructionSourceRunIndex < 0)
            {
                field.InstructionSourceRunIndex = currentSourceRunIndex;
            }

            field.Instruction.Append(instructionText);
            field.InstructionRunCount++;
            string? fieldPlaceholder = ResolveFieldPlaceholder(field.Instruction.ToString());
            if (fieldPlaceholder is not null && !field.PlaceholderEmitted)
            {
                field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
                int runsBefore = runs.Count;
                AddResolvedTextRuns(
                    runs,
                    resolvedRun.AllCaps == true ? fieldPlaceholder.ToUpperInvariant() : fieldPlaceholder,
                    resolvedRun,
                    runStyleResolution,
                    currentSourceRunIndex,
                    sourceTextOffsetInRun: 0,
                    revision,
                    revisions);
                field.ResultRunCount += runs.Count - runsBefore;
                field.PlaceholderEmitted = true;
                currentPageInstructionSeen = true;
            }
        }

        void AddComplexFieldText(
            string text,
            int currentSourceRunIndex,
            int sourceTextOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            DocxComplexFieldState[] resultFields = complexFieldStack
                .Where(field => field.InResult)
                .ToArray();
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.HasCachedResult = true;
            }

            string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
            bool suppressPageNumberCache = currentPageInstructionSeen && displayText.Trim().All(char.IsDigit);
            if (suppressPageNumberCache)
            {
                return;
            }

            if (resultFields.Length != 0)
            {
                int fieldTextRunIndex = runs.Count;
                int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                foreach (DocxComplexFieldState field in resultFields)
                {
                    field.EnsureTextSpan(fieldTextRunIndex, fieldTextLengthStart);
                }
            }

            int runsBefore = runs.Count;
            AddResolvedTextRuns(
                runs,
                displayText,
                resolvedRun,
                runStyleResolution,
                currentSourceRunIndex,
                sourceTextOffset,
                revision,
                revisions);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
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

        void AddOrderedRunTextAndReferences(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (IsInlineReferenceElement(child))
                {
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions);
                }
                else
                {
                    string childText = ReadRunTextChild(child);
                    if (childText.Length != 0)
                    {
                        DocxComplexFieldState[] resultFields = ActiveComplexResultFields();
                        if (resultFields.Length != 0)
                        {
                            int fieldTextRunIndex = runs.Count;
                            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                            foreach (DocxComplexFieldState field in resultFields)
                            {
                                field.HasCachedResult = true;
                                field.EnsureTextSpan(fieldTextRunIndex, fieldTextLengthStart);
                            }
                        }

                        int runsBefore = runs.Count;
                        AddResolvedTextRuns(
                            runs,
                            resolvedRun.AllCaps == true ? childText.ToUpperInvariant() : childText,
                            resolvedRun,
                            runStyleResolution,
                            currentSourceRunIndex,
                            textOffset,
                            revision,
                            revisions);
                        int addedRuns = runs.Count - runsBefore;
                        foreach (DocxComplexFieldState field in resultFields)
                        {
                            field.RendersCachedResult = true;
                            field.ResultRunCount += addedRuns;
                        }
                    }

                    textOffset += childText.Length;
                }

                childIndex++;
            }
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
                AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRuns, revision, revisions);

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        void AddInlineReference(
            XElement child,
            int currentSourceRunIndex,
            int childIndex,
            int textOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            bool emitDisplayRun,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            if (ResolveInlineReferenceKind(child) is not { } kind)
            {
                return;
            }

            string? customMarkFollows = kind == DocxRelatedStoryKind.Footnote || kind == DocxRelatedStoryKind.Endnote
                ? (string?)child.Attribute(WordprocessingNamespace + "customMarkFollows")
                : null;
            string? displayText = ResolveInlineReferenceDisplayText(kind, customMarkFollows);
            inlineReferences.Add(new DocxInlineReference(
                kind,
                (string?)child.Attribute(WordprocessingNamespace + "id"),
                customMarkFollows,
                displayText,
                currentSourceRunIndex,
                childIndex,
                textOffset)
            {
                Revision = revision,
                Revisions = revisions
            });
            if (kind == DocxRelatedStoryKind.Comment)
            {
                AddCommentReferenceRange((string?)child.Attribute(WordprocessingNamespace + "id"), currentSourceRunIndex, textOffset);
            }

            if (emitDisplayRun && displayText is not null)
            {
                AddInlineReferenceDisplayRun(displayText, currentSourceRunIndex, textOffset, resolvedRun, runStyleResolution, revision);
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
                    textOffset));
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
            if (!string.IsNullOrEmpty(customMarkFollows) || (kind != DocxRelatedStoryKind.Footnote && kind != DocxRelatedStoryKind.Endnote))
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
            inlineReferenceCounters.TryGetValue(kind, out int current);
            int next = current == 0 ? settings.NumberStart ?? 1 : current + 1;
            inlineReferenceCounters[kind] = next;
            return FormatNoteReferenceNumber(next, settings.NumberFormatValue);
        }

        void AddInlineReferenceDisplayRun(
            string displayText,
            int currentSourceRunIndex,
            int textOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision)
        {
            AddResolvedTextRuns(
                runs,
                displayText,
                resolvedRun with { VerticalAlignmentValue = "superscript" },
                runStyleResolution,
                currentSourceRunIndex,
                textOffset,
                revision,
                revisions: null);
        }

        static bool IsInlineReferenceElement(XElement element)
        {
            return ResolveInlineReferenceKind(element) is not null;
        }

        static DocxRelatedStoryKind? ResolveInlineReferenceKind(XElement element)
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
    }

    private static void AddBookmarkAnchor(
        XElement bookmarkStart,
        List<DocxBookmarkAnchor> bookmarkAnchors,
        int sourceRunIndex,
        List<DocxTextRun> runs)
    {
        bookmarkAnchors.Add(new DocxBookmarkAnchor(
            (string?)bookmarkStart.Attribute(WordprocessingNamespace + "id"),
            (string?)bookmarkStart.Attribute(WordprocessingNamespace + "name"),
            sourceRunIndex,
            runs.Count,
            runs.Sum(run => run.Text.Length)));
    }

    private static void AddCommentRangeStart(
        XElement rangeStart,
        List<DocxCommentRangeStart> openCommentRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs)
    {
        openCommentRanges.Add(new DocxCommentRangeStart(
            (string?)rangeStart.Attribute(WordprocessingNamespace + "id"),
            sourceRunIndex,
            runs.Sum(run => run.Text.Length)));
    }

    private static void AddCommentRangeEnd(
        XElement rangeEnd,
        List<DocxCommentRangeStart> openCommentRanges,
        List<DocxCommentRange> commentRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs)
    {
        string? id = (string?)rangeEnd.Attribute(WordprocessingNamespace + "id");
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
            runs.Sum(run => run.Text.Length),
            ReferenceSourceRunIndex: null,
            ReferenceTextOffset: null));
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static void AddResolvedTextRuns(
        List<DocxTextRun> runs,
        string text,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution,
        int sourceRunIndex,
        int sourceTextOffsetInRun,
        DocxRevisionInfo? revision,
        IReadOnlyList<DocxRevisionInfo>? revisions)
    {
        var segment = new StringBuilder();
        bool? currentComplexScript = null;
        int segmentSourceOffset = sourceTextOffsetInRun;
        foreach (Rune rune in text.EnumerateRunes())
        {
            bool complexScript = DocxScriptClassifier.IsComplexScriptRune(rune.Value);
            if (currentComplexScript is not null && currentComplexScript.Value != complexScript)
            {
                string segmentText = segment.ToString();
                AddResolvedTextRun(runs, segmentText, resolvedRun, styleResolution, currentComplexScript.Value, sourceRunIndex, segmentSourceOffset, revision, revisions);
                segmentSourceOffset += segmentText.Length;
                segment.Clear();
            }

            segment.Append(rune);
            currentComplexScript = complexScript;
        }

        if (segment.Length != 0 && currentComplexScript is not null)
        {
            AddResolvedTextRun(runs, segment.ToString(), resolvedRun, styleResolution, currentComplexScript.Value, sourceRunIndex, segmentSourceOffset, revision, revisions);
        }
    }

    private static void AddResolvedTextRun(
        List<DocxTextRun> runs,
        string text,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution,
        bool complexScript,
        int sourceRunIndex,
        int sourceTextOffsetInRun,
        DocxRevisionInfo? revision,
        IReadOnlyList<DocxRevisionInfo>? revisions)
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
        runs.Add(new DocxTextRun(
            text,
            resolvedRun.FontSize ?? DocxDefaults.FontSizePoints,
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
            StyleResolution = styleResolution,
            SourceRunIndex = sourceRunIndex,
            SourceTextOffsetInRun = sourceTextOffsetInRun,
            Revision = revision,
            Revisions = revisions ?? RevisionList(revision)
        });
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

    private static string ReadRunTextChild(XElement child)
    {
        if (child.Name == WordprocessingNamespace + "t" ||
            child.Name == WordprocessingNamespace + "delText")
        {
            return (string?)child ?? string.Empty;
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

        if (child.Name == WordprocessingNamespace + "cr")
        {
            return "\n";
        }

        if (child.Name == WordprocessingNamespace + "br" &&
            string.IsNullOrEmpty((string?)child.Attribute(WordprocessingNamespace + "type")))
        {
            return "\n";
        }

        return string.Empty;
    }

    private static string? ReadParagraphStyleId(XElement? paragraphProperties)
    {
        return (string?)paragraphProperties?
            .Element(WordprocessingNamespace + "pStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }
}
