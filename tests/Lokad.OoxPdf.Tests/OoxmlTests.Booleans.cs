using System.Xml;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxmlTests
{
    public static void XmlBooleanWhitespacePreservesSchemaValues()
    {
        foreach (string token in new[] { "true", "false", "1", "0" })
        foreach (string padding in new[] { "", " ", "\t", "\r\n", " \t\r\n" })
        {
            string value = padding + token + padding;
            bool expected = XmlConvert.ToBoolean(value);
            var element = new XElement("flag", new XAttribute("val", value));
            TestAssert.Equal(expected, OoxBoolean.IsTrue(value));
            TestAssert.Equal(!expected, OoxBoolean.IsOff(value));
            TestAssert.Equal(expected, OoxXml.ReadBool(element, "val"));
            TestAssert.Equal(expected, OoxBoolean.ParseAttribute(element, "val", !expected));
            TestAssert.Equal((bool?)expected, OoxBoolean.ParseOptionalAttribute(element, "val"));
            TestAssert.Equal(expected, OoxBoolean.ParseElement(element, !expected, null));
        }
        TestAssert.True(OoxBoolean.IsTrue(" \tTrUe\r\n"), "Existing case-insensitive true values retain their meaning.");
        TestAssert.True(OoxBoolean.IsOff(" \tFaLsE\r\n"), "Existing case-insensitive false values retain their meaning.");
    }

    public static void BooleanReadersPreserveVocabularyUnknownsAndMissingValues()
    {
        TestAssert.True(OoxBoolean.IsTrue("ON") && OoxBoolean.IsOff("OFF"), "Raw legacy on/off values remain supported.");
        TestAssert.True(!OoxXml.ReadBool(new XElement("flag", new XAttribute("val", "on")), "val"),
            "The strict DrawingML reader does not acquire the on/off vocabulary.");
        foreach (string? value in new string?[] { null, "", " ", "unknown", " on ", " off ", "\u00A0true\u00A0", "\u00A0false\u00A0", "true false", "10" })
        {
            TestAssert.True(!OoxBoolean.IsTrue(value), "Unknown values do not become true.");
            TestAssert.True(!OoxBoolean.IsOff(value), "Unknown display values retain the shown fallback.");
        }
        var missing = new XElement("flag");
        TestAssert.Equal(true, OoxBoolean.ParseAttribute(missing, "val", true));
        TestAssert.Equal(false, OoxBoolean.ParseAttribute(missing, "val", false));
        TestAssert.Equal((bool?)null, OoxBoolean.ParseOptionalAttribute(missing, "val"));
        TestAssert.Equal(true, OoxBoolean.ParseElement(missing, false, null));
        TestAssert.Equal(false, OoxBoolean.ParseElement(null, false, null));
        TestAssert.Equal(true, OoxBoolean.ParseElement(null, true, null));
        TestAssert.Equal((bool?)false, OoxBoolean.ParseOptionalAttribute(new XElement("flag", new XAttribute("val", "unknown")), "val"));
    }
}
