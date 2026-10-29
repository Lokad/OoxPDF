using System.Globalization;
using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Ooxml;

// Adapt common Strict DOCX vocabulary to the existing reader. This is deliberately
// not a general Strict converter: unknown namespaces and property forms survive
// unchanged, and the reader retains the dialect warning.
internal static class OoxStrictDocx
{
    private sealed class StrictOrigin { }
    private static readonly StrictOrigin Origin = new();

    public static bool HasNormalizedRoot(XDocument document) => document.Annotation<StrictOrigin>() is not null;

    public static void Normalize(XDocument document, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (document.Root is not { } root || Alias(root.Name.NamespaceName) is null)
            return;

        document.AddAnnotation(Origin);
        foreach (XElement element in root.DescendantsAndSelf())
        {
            cancellationToken.ThrowIfCancellationRequested();
            element.Name = NormalizeName(element.Name);
            foreach (XAttribute attribute in element.Attributes().ToArray())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (attribute.IsNamespaceDeclaration)
                {
                    if (Alias(attribute.Value) is { } target) attribute.Value = target.NamespaceName;
                    continue;
                }

                XName name = NormalizeName(attribute.Name);
                string value = NormalizeMeasure(element, name, attribute.Value);
                if (name == attribute.Name)
                    attribute.Value = value;
                else
                {
                    if (element.Attribute(name) is not null)
                        throw new InvalidDataException("Conflicting Strict and transitional XML attributes.");
                    attribute.Remove();
                    element.Add(new XAttribute(name, value));
                }
            }
        }
    }

    private static XName NormalizeName(XName name) => Alias(name.NamespaceName) is { } target
        ? target + name.LocalName : name;

    private static XNamespace? Alias(string value) => value switch
    {
        "http://purl.oclc.org/ooxml/wordprocessingml/main" => WordprocessingNamespace,
        "http://purl.oclc.org/ooxml/drawingml/main" => DrawingNamespace,
        "http://purl.oclc.org/ooxml/drawingml/picture" => "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" => WordprocessingDrawingNamespace,
        "http://purl.oclc.org/ooxml/officeDocument/math" => MathNamespace,
        "http://purl.oclc.org/ooxml/officeDocument/relationships" => RelationshipsNamespace,
        _ => null
    };

    private static string NormalizeMeasure(XElement element, XName name, string value)
    {
        if (element.Name.Namespace == DrawingNamespace && name.Namespace == XNamespace.None && value.EndsWith('%') &&
            IsDrawingPercentage(element.Name.LocalName, name.LocalName))
            return Scale(value.AsSpan(0, value.Length - 1), 1000m);
        if (element.Name.Namespace != WordprocessingNamespace || name.Namespace != WordprocessingNamespace)
            return value;

        string local = element.Name.LocalName;
        string attribute = name.LocalName;
        if (value.EndsWith('%'))
        {
            if (attribute == "w" && IsTableWidth(local))
                return Scale(value.AsSpan(0, value.Length - 1), 50m);
            if (local == "zoom" && attribute == "percent")
                return Scale(value.AsSpan(0, value.Length - 1), 1m);
            return value;
        }
        if (value.Length < 3) return value;

        decimal? unitsPerPoint = (local, attribute) switch
        {
            ("sz" or "szCs" or "kern" or "position" or "size" or "hps" or "hpsBaseText", "val") => 2m,
            ("pgSz", "w" or "h") => 20m,
            ("pgMar", "top" or "right" or "bottom" or "left" or "header" or "footer" or "gutter") => 20m,
            ("cols" or "col", "w" or "space") => 20m,
            ("spacing", "before" or "after" or "line" or "val") => 20m,
            ("ind", "left" or "right" or "start" or "end" or "hanging" or "firstLine") => 20m,
            ("tab", "pos") => 20m,
            ("defaultTabStop" or "trHeight" or "fitText" or "hyphenationZone", "val") => 20m,
            ("framePr", "w" or "h" or "x" or "y" or "hSpace" or "vSpace") => 20m,
            (_, "w") when IsTableWidth(local) => 20m,
            _ => null
        };
        if (unitsPerPoint is null) return value;
        decimal? pointsPerUnit = value.AsSpan(value.Length - 2) switch
        {
            "pt" => 1m,
            "in" => 72m,
            "cm" => 72m / 2.54m,
            "mm" => 72m / 25.4m,
            "pc" or "pi" => 12m,
            _ => null
        };
        return pointsPerUnit is { } factor
            ? Scale(value.AsSpan(0, value.Length - 2), factor * unitsPerPoint.Value) : value;
    }

    private static bool IsTableWidth(string element) => element is
        "tblW" or "tcW" or "tblInd" or "tblCellSpacing" or "wBefore" or "wAfter" or
        "top" or "bottom" or "left" or "right" or "start" or "end";

    private static bool IsDrawingPercentage(string element, string attribute) => (element, attribute) switch
    {
        ("gs", "pos") or ("miter", "lim") => true,
        ("srcRect" or "fillRect", "l" or "r" or "t" or "b") => true,
        ("normAutofit", "fontScale" or "lnSpcReduction") => true,
        ("alpha" or "alphaMod" or "alphaOff" or "tint" or "shade" or
            "lum" or "lumMod" or "lumOff" or "sat" or "satMod" or "satOff" or
            "hueMod" or "red" or "redMod" or "redOff" or "green" or "greenMod" or "greenOff" or
            "blue" or "blueMod" or "blueOff", "val") => true,
        _ => false
    };

    private static string Scale(ReadOnlySpan<char> number, decimal factor)
    {
        if (!decimal.TryParse(number, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out decimal value))
            throw new InvalidDataException("Malformed Strict OOXML measure.");
        try
        {
            long result = checked((long)decimal.Round(value * factor, 0, MidpointRounding.AwayFromZero));
            return result.ToString(CultureInfo.InvariantCulture);
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("Strict OOXML measure is outside the supported range.", ex);
        }
    }
}
