using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal static class OoxBoolean
{
    public static bool IsTrue(string? value)
    {
        return IsXmlTrue(value) || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
    }

    // XML booleans collapse XML whitespace, not arbitrary Unicode spaces.
    // Keep the legacy on/off vocabulary separate from this strict boolean reader.
    public static bool IsXmlTrue(string? value)
    {
        ReadOnlySpan<char> token = value.AsSpan().Trim(" \t\r\n".AsSpan());
        return token.SequenceEqual("1") || token.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsXmlFalse(string? value, StringComparison comparison)
    {
        ReadOnlySpan<char> token = value.AsSpan().Trim(" \t\r\n".AsSpan());
        return token.SequenceEqual("0") || token.Equals("false", comparison);
    }

    public static bool ParseAttribute(XElement? element, XName attributeName, bool defaultValue)
    {
        string? value = (string?)element?.Attribute(attributeName);
        return value is null ? defaultValue : IsTrue(value);
    }

    public static bool? ParseOptionalAttribute(XElement? element, XName attributeName)
    {
        XAttribute? attribute = element?.Attribute(attributeName);
        return attribute is null ? null : IsTrue(attribute.Value);
    }

    // Explicit off-values for display flags; unknown or absent values fail
    // open (shown), so garbage never drops content (S01).
    public static bool IsOff(string? value) =>
        IsXmlFalse(value, StringComparison.OrdinalIgnoreCase) || string.Equals(value, "off", StringComparison.OrdinalIgnoreCase);

    public static bool ParseElement(XElement? element, bool defaultValue, XName? valueAttributeName)
    {
        if (element is null)
        {
            return defaultValue;
        }

        string? value = (string?)element.Attribute(valueAttributeName ?? "val");
        return value is null || IsTrue(value);
    }

}
