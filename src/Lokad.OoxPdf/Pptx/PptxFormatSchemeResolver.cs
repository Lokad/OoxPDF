using System.Globalization;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxFormatSchemeResolver
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static PptxFormatSchemeReference ResolveFillReference(XElement shape, PptxTheme theme)
    {
        XElement? reference = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fillRef");
        int index = ReadIndex(reference);
        return index > 0 && theme.TryGetFillStyle(index, out XElement style)
            ? new PptxFormatSchemeReference(reference, index, style)
            : new PptxFormatSchemeReference(reference, index, null);
    }

    public static PptxFormatSchemeReference ResolveLineReference(XElement shape, PptxTheme theme)
    {
        XElement? reference = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "lnRef");
        int index = ReadIndex(reference);
        return index > 0 && theme.TryGetLineStyle(index, out XElement style)
            ? new PptxFormatSchemeReference(reference, index, style)
            : new PptxFormatSchemeReference(reference, index, null);
    }

    public static int ReadIndex(XElement? reference)
    {
        return reference?.Attribute("idx") is { } attribute &&
            int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
            ? index
            : 0;
    }
}

internal readonly record struct PptxFormatSchemeReference(
    XElement? Reference,
    int Index,
    XElement? Style);
