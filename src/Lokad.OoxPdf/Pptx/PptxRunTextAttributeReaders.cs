using System.Globalization;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

// R14: one shared run-attribute interpretation for scene building and rendering.
// These seven readers were character-identical private copies in PptxSceneBuilder
// and PptxRenderer; both sides now call this single implementation, so run-level
// text interpretation cannot drift between layout and inspection. Raw source stays
// provenance: callers keep the original elements for everything else.
internal static class PptxRunTextAttributeReaders
{
    internal static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    internal static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    internal static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    internal static bool IsStrikeEnabled(string? value)
    {
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    internal static string? ReadUnderlineValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"));
    }

    internal static string? ReadStrikeValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
    }

    internal static bool HasRunTextFill(XElement? runProperties)
    {
        if (runProperties is null)
        {
            return false;
        }

        XNamespace ns = runProperties.Name.Namespace;
        return runProperties.Element(ns + "solidFill") is not null ||
            runProperties.Element(ns + "noFill") is not null ||
            runProperties.Element(ns + "gradFill") is not null;
    }
    internal static string? ReadTextCapsValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
    }

}
