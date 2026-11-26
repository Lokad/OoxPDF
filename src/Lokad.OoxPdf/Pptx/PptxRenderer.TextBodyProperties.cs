using System.Globalization;
using System.Xml.Linq;

using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static PptxTextBodyProperties ReadTextBodyProperties(XElement textBody, XElement? inheritedTextBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        XElement? inheritedBodyProperties = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        (string? columnCountValue, PptxTextBodyPropertySource columnCountSource) =
            ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "numCol");
        (string? columnSpacingValue, PptxTextBodyPropertySource columnSpacingSource) =
            ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "spcCol");
        int columnCount = int.TryParse(columnCountValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count)
            ? Math.Clamp(count, 1, 16) : 1;
        double columnSpacing = long.TryParse(columnSpacingValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long spacing)
            ? Math.Max(0d, OoxUnits.EmuToPoints(spacing)) : 0d;
        PptxTextBodyPropertySource columnSource = MergeColumnSources(columnCountSource, columnSpacingSource);
        (TextInsets insets, TextInsetSources insetSources, TextInsetValues insetValues) = ReadTextInsets(bodyProperties, inheritedBodyProperties);
        (XElement? autofit, string autofitMode, PptxTextBodyPropertySource autofitModeSource) = ReadTextAutofit(bodyProperties, inheritedBodyProperties);
        (double fontScale, PptxTextBodyPropertySource fontScaleSource, string? fontScaleValue) = ReadNormAutofitFontScale(autofit, autofitModeSource);
        (double lineSpacingScale, PptxTextBodyPropertySource lineSpacingScaleSource, string? lineSpacingReductionValue) = ReadNormAutofitLineSpacingScale(autofit, autofitModeSource);
        (string? orientation, PptxTextBodyPropertySource orientationSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "vert");
        (string? verticalAnchor, PptxTextBodyPropertySource verticalAnchorSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "anchor");
        (string? anchorCenter, PptxTextBodyPropertySource anchorCenterSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "anchorCtr");
        (string? wrap, PptxTextBodyPropertySource wrapSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "wrap");
        (string? verticalOverflow, PptxTextBodyPropertySource verticalOverflowSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "vertOverflow");
        (string? horizontalOverflow, _) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "horzOverflow");
        (string? compatibleLineSpacing, PptxTextBodyPropertySource compatibleLineSpacingSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "compatLnSpc");
        (string? rotation, PptxTextBodyPropertySource rotationSource) = ReadTextBodyAttributeWithSource(bodyProperties, inheritedBodyProperties, "rot");
        return new PptxTextBodyProperties(
            insets,
            insetSources,
            insetValues,
            ParseTextOrientation(orientation),
            orientation,
            orientationSource,
            ParseTextVerticalAnchor(verticalAnchor),
            verticalAnchor,
            verticalAnchorSource,
            anchorCenter is null ? null : OoxBoolean.IsTrue(anchorCenter),
            anchorCenter,
            anchorCenterSource,
            ParseTextWrapMode(wrap),
            wrap,
            wrapSource,
            ParseTextVerticalOverflow(verticalOverflow),
            verticalOverflow,
            verticalOverflowSource,
            columnCount,
            columnSpacing,
            columnSource,
            columnCountSource,
            columnSpacingSource,
            columnCountValue,
            columnSpacingValue,
            autofitMode,
            autofitModeSource,
            fontScale,
            fontScaleValue,
            fontScaleSource,
            lineSpacingScale,
            lineSpacingReductionValue,
            lineSpacingScaleSource,
            compatibleLineSpacing is not null && OoxBoolean.IsTrue(compatibleLineSpacing),
            compatibleLineSpacing,
            compatibleLineSpacingSource,
            ParseTextBodyRotationDegrees(rotation),
            rotation,
            rotationSource)
        {
            AutofitMode = ParseTextAutofitMode(autofitMode),
            ClipsHorizontalOverflow = string.Equals(horizontalOverflow, "clip", StringComparison.OrdinalIgnoreCase) ? true :
                string.Equals(horizontalOverflow, "overflow", StringComparison.OrdinalIgnoreCase) ? false : null
        };
    }

    private static PptxTextBodyPropertySource MergeColumnSources(PptxTextBodyPropertySource count, PptxTextBodyPropertySource spacing)
    {
        if (count == PptxTextBodyPropertySource.DirectBodyPr || spacing == PptxTextBodyPropertySource.DirectBodyPr)
        {
            return PptxTextBodyPropertySource.DirectBodyPr;
        }

        return count == PptxTextBodyPropertySource.InheritedBodyPr || spacing == PptxTextBodyPropertySource.InheritedBodyPr
            ? PptxTextBodyPropertySource.InheritedBodyPr
            : PptxTextBodyPropertySource.DefaultValue;
    }

    private static (string? Value, PptxTextBodyPropertySource Source) ReadTextBodyAttributeWithSource(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties,
        string attributeName)
    {
        if (bodyProperties?.Attribute(attributeName) is { } directAttribute)
        {
            return (directAttribute.Value, PptxTextBodyPropertySource.DirectBodyPr);
        }

        if (inheritedBodyProperties?.Attribute(attributeName) is { } inheritedAttribute)
        {
            return (inheritedAttribute.Value, PptxTextBodyPropertySource.InheritedBodyPr);
        }

        return (null, PptxTextBodyPropertySource.DefaultValue);
    }

    private static TextInsets ReadPresetTextRectInsets(XElement shape, double width, double height)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        string? preset = shapeProperties
            ?.Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst")
            ?.Value;

        return preset switch
        {
            "ellipse" => new TextInsets(
                width * PptxTextMetricRules.EllipseTextRectInsetRatio,
                width * PptxTextMetricRules.EllipseTextRectInsetRatio,
                height * PptxTextMetricRules.EllipseTextRectInsetRatio,
                height * PptxTextMetricRules.EllipseTextRectInsetRatio),
            "roundRect" when shapeProperties is not null => RoundRectTextRectInsets(shapeProperties, width, height),
            _ => TextInsets.Empty
        };
    }

    private static TextInsets RoundRectTextRectInsets(XElement shapeProperties, double width, double height)
    {
        // Text insets historically read the first matching declaration.
        // Paint uses the scene's adjustment map.
        string? formula = (string?)shapeProperties.Element(DrawingNamespace + "prstGeom")
            ?.Element(DrawingNamespace + "avLst")
            ?.Elements(DrawingNamespace + "gd")
            .FirstOrDefault(element => (string?)element.Attribute("name") == "adj")?.Attribute("fmla");
        double adjustment = formula is not null && formula.StartsWith("val ", StringComparison.Ordinal) &&
            double.TryParse(formula.AsSpan(4), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value
                : PptxTextMetricRules.RoundRectDefaultAdjustment;
        adjustment = Math.Clamp(adjustment, 0d, 50000d) / 100000d;
        double radius = Math.Min(width, height) * adjustment;
        double inset = radius * PptxTextMetricRules.RoundRectTextRectRadiusInsetFactor;
        return new TextInsets(inset, inset, inset, inset);
    }
}
