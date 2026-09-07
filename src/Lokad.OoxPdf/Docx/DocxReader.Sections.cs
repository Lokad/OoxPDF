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
    private static DocxSectionBreakElement ReadSectionBreak(
        XElement sectionProperties,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        XDocument? settings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        DocxRevisionInfo? inheritedRevision)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? columns = sectionProperties.Element(WordprocessingNamespace + "cols");
        return new DocxSectionBreakElement(
            ReadPageSettings(
                sectionProperties.Element(WordprocessingNamespace + "pgSz"),
                sectionProperties.Element(WordprocessingNamespace + "pgMar"),
                sectionProperties,
                settings,
                package,
                relationships,
                styles,
                numbering,
                markupMode,
                cancellationToken),
            DocxSectionBreakTypeExtensions.FromValue((string?)sectionProperties
                .Element(WordprocessingNamespace + "type")
                ?.Attribute(WordprocessingNamespace + "val")),
            (string?)columns?.Attribute(WordprocessingNamespace + "num"),
            (string?)columns?.Attribute(WordprocessingNamespace + "equalWidth"),
            (string?)columns?.Attribute(WordprocessingNamespace + "space"),
            ReadSectionColumns(columns))
        {
            Revisions = MergeRevisionLists(inheritedRevision, ReadPropertyChangeRevisions(sectionProperties))
        };
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxSectionColumn> ReadSectionColumns(XElement? columns)
    {
        return columns
            ?.Elements(WordprocessingNamespace + "col")
            .Select(column => new DocxSectionColumn(
                (string?)column.Attribute(WordprocessingNamespace + "w"),
                (string?)column.Attribute(WordprocessingNamespace + "space")))
            .ToArray() ?? [];
    }

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XElement story,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement[] paragraphs = story
            .Descendants(WordprocessingNamespace + "p")
            .ToArray();
        XElement[] bodyBlocks = story
            .Elements()
            .Where(IsBodyBlockElement)
            .ToArray();
        var drawings = new List<DocxFloatingDrawing>();
        foreach (XElement anchor in story.Descendants(WordprocessingDrawingNamespace + "anchor"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRevisionInfo? revision = FindInheritedRevision(anchor, markupMode);
            if (IsInsideExcludedRevisionContainer(anchor, markupMode))
            {
                continue;
            }

            drawings.Add(ReadFloatingDrawing(
                anchor,
                package,
                relationships,
                styles,
                numbering,
                FindSourceParagraphIndex(anchor, paragraphs),
                FindSourceBlockIndex(anchor, bodyBlocks),
                revision,
                markupMode,
                cancellationToken));
        }

        return drawings;
    }

    private static IReadOnlyList<DocxBodyElement> ReadRelatedStoryBodyElements(
        IEnumerable<XElement> elements,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        var bodyElements = new List<DocxBodyElement>();
        foreach (DocxRevisionScopedElement scopedElement in EnumerateRevisionScopedChildren(elements, markupMode, WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = scopedElement.Element;
            DocxRevisionInfo? inheritedRevision = scopedElement.Revision;
            if (element.Name == WordprocessingNamespace + "p")
            {
                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: null, documentSettings: null, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                if (paragraph is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: null, documentSettings: null, markupMode: markupMode, cancellationToken: cancellationToken, inheritedRevision: inheritedRevision);
                if (table is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateTable(table));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(bodyElements, markupMode);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxInlineImage> ReadInlineImages(
        XElement run,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        var images = new List<DocxInlineImage>();
        foreach (XElement inline in run.Descendants(WordprocessingDrawingNamespace + "inline"))
        {
            DocxInlineImage? image = ReadDrawingImage(inline, package, relationships, revision);
            if (image is null)
            {
                continue;
            }

            images.Add(image);
        }

        foreach (XElement shape in run.Descendants(VmlNamespace + "shape"))
        {
            DocxInlineImage? image = ReadVmlInlineImage(shape, package, relationships, revision);
            if (image is null)
            {
                continue;
            }

            images.Add(image);
        }

        return images;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    // Inline DrawingML textboxes (wp:inline plus wps:txbx, no blip) parse beside inline
    // pictures; the content joins body flow as a block at layout time.
    private static IReadOnlyList<DocxInlineTextBox> ReadInlineTextBoxes(
        XElement run,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        DocxRevisionInfo? revision,
        CancellationToken cancellationToken)
    {
        var textBoxes = new List<DocxInlineTextBox>();
        foreach (XElement inline in run.Descendants(WordprocessingDrawingNamespace + "inline"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ReadDrawingImage(inline, package, relationships, revision) is not null)
            {
                continue;
            }

            XElement? textBoxContent = inline
                .Descendants(WordprocessingNamespace + "txbxContent")
                .FirstOrDefault();
            if (textBoxContent is null)
            {
                continue;
            }

            XElement? extent = inline.Element(WordprocessingDrawingNamespace + "extent");
            XElement? textBoxBodyProperties = inline
                .Descendants(WordprocessingShapeNamespace + "bodyPr")
                .FirstOrDefault();
            textBoxes.Add(new DocxInlineTextBox(
                (string?)extent?.Attribute("cx"),
                (string?)extent?.Attribute("cy"),
                (string?)textBoxBodyProperties?.Attribute("lIns"),
                (string?)textBoxBodyProperties?.Attribute("tIns"),
                (string?)textBoxBodyProperties?.Attribute("rIns"),
                (string?)textBoxBodyProperties?.Attribute("bIns"))
            {
                Revisions = RevisionList(revision),
                BodyElements = ReadRelatedStoryBodyElements(
                    textBoxContent.Elements(),
                    styles,
                    numbering,
                    new Dictionary<(string NumId, int Level), int>(),
                    package,
                    relationships,
                    markupMode,
                    cancellationToken)
            });
        }

        return textBoxes;
    }

    private static DocxInlineImage? ReadDrawingImage(
        XElement drawing,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        XElement? extent = drawing.Element(WordprocessingDrawingNamespace + "extent");
        string? relationshipId = ReadDrawingImageRelationshipId(drawing);
        if (extent is null || relationshipId is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            return null;
        }

        return new DocxInlineImage(
            OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(extent, "cx", "DOCX")),
            OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(extent, "cy", "DOCX")),
            imagePart.ContentType,
            imagePart.Bytes,
            imagePart.Name)
        {
            Revisions = RevisionList(revision)
        };
    }

    private static string? ReadDrawingImageRelationshipId(XElement drawing)
    {
        return (string?)drawing
            .Descendants(DrawingNamespace + "blip")
            .FirstOrDefault()
            ?.Attribute(RelationshipsNamespace + "embed");
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxInlineImage? ReadVmlInlineImage(
        XElement shape,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        if (!TryReadVmlImageShape(
                shape,
                relationships,
                out OoxRelationship? relationship,
                out double widthPoints,
                out double heightPoints) ||
            relationship is null ||
            relationship.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            return null;
        }

        return new DocxInlineImage(
            widthPoints,
            heightPoints,
            imagePart.ContentType,
            imagePart.Bytes,
            imagePart.Name)
        {
            Revisions = RevisionList(revision)
        };
    }

    private static bool TryReadVmlImageShape(
        XElement shape,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        out OoxRelationship? relationship,
        out double widthPoints,
        out double heightPoints)
    {
        relationship = null;
        widthPoints = 0d;
        heightPoints = 0d;
        XElement? imageData = shape.Descendants(VmlNamespace + "imagedata").FirstOrDefault();
        string? relationshipId = (string?)imageData?.Attribute(RelationshipsNamespace + "id");
        if (relationshipId is null ||
            !relationships.TryGetValue(relationshipId, out OoxRelationship? resolvedRelationship) ||
            resolvedRelationship.IsExternal ||
            resolvedRelationship.ResolvedTarget is null ||
            !TryReadVmlShapeSizePoints(shape, out widthPoints, out heightPoints))
        {
            return false;
        }

        relationship = resolvedRelationship;
        return true;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool TryReadVmlShapeSizePoints(XElement shape, out double widthPoints, out double heightPoints)
    {
        widthPoints = 0d;
        heightPoints = 0d;
        double? width = null;
        double? height = null;
        string? style = (string?)shape.Attribute("style");
        if (!string.IsNullOrWhiteSpace(style))
        {
            foreach (string declaration in style.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                int separator = declaration.IndexOf(':');
                if (separator <= 0 || separator == declaration.Length - 1)
                {
                    continue;
                }

                string name = declaration[..separator].Trim();
                string value = declaration[(separator + 1)..].Trim();
                if (name.Equals("width", StringComparison.OrdinalIgnoreCase) &&
                    TryParseVmlLengthPoints(value, out double parsedWidth))
                {
                    width = parsedWidth;
                }
                else if (name.Equals("height", StringComparison.OrdinalIgnoreCase) &&
                    TryParseVmlLengthPoints(value, out double parsedHeight))
                {
                    height = parsedHeight;
                }
            }
        }

        if (width is null &&
            shape.Attribute("width") is { } widthAttribute &&
            TryParseVmlLengthPoints(widthAttribute.Value, out double attributeWidth))
        {
            width = attributeWidth;
        }

        if (height is null &&
            shape.Attribute("height") is { } heightAttribute &&
            TryParseVmlLengthPoints(heightAttribute.Value, out double attributeHeight))
        {
            height = attributeHeight;
        }

        if (width is not > 0d || height is not > 0d)
        {
            return false;
        }

        widthPoints = width.Value;
        heightPoints = height.Value;
        return true;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool TryParseVmlLengthPoints(string value, out double points)
    {
        points = 0d;
        string trimmed = value.Trim().ToLowerInvariant();
        if (trimmed.Length == 0)
        {
            return false;
        }

        int unitStart = trimmed.Length;
        while (unitStart > 0 && char.IsLetter(trimmed[unitStart - 1]))
        {
            unitStart--;
        }

        string numberText = trimmed[..unitStart].Trim();
        string unit = trimmed[unitStart..];
        if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric))
        {
            return false;
        }

        points = unit switch
        {
            "" or "pt" => numeric,
            "in" => numeric * 72d,
            "cm" => numeric * 72d / 2.54d,
            "mm" => numeric * 72d / 25.4d,
            "pc" => numeric * 12d,
            "px" => numeric * 0.75d,
            _ => 0d
        };
        return points > 0d;
    }
}
