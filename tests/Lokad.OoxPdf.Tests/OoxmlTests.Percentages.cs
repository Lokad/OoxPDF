using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxmlTests
{
    public static void OoxPercentageFormsKeepSignedUnitsAndFiniteValues()
    {
        foreach (var (integer, literal, expected) in new[]
        {
            ("0", "0%", 0d), ("30000", "30%", .3d), ("-12500", "-12.5%", -.125d),
            ("1", "0.001%", .00001d), ("38970", "38.97%", .3897d), ("125000", "+125%", 1.25d),
            (" 50000 ", " 50% ", .5d)
        })
        {
            foreach (string source in new[] { integer, literal })
            {
                TestAssert.True(OoxXml.TryParsePercentage(source, out double actual), "A supported percentage must parse.");
                TestAssert.Equal(expected, actual);
                TestAssert.Equal(expected, OoxXml.ParsePercentage(source));
            }
        }
        foreach (string source in new[] { "", "bad", "30%%", "NaN%", "Infinity%", "1e2%", "1,5%", "30 %" })
        {
            TestAssert.True(!OoxXml.TryParsePercentage(source, out _), "Malformed percentage must not acquire a value: " + source);
            TestAssert.Throws<FormatException>(() => OoxXml.ParsePercentage(source));
        }
        TestAssert.True(!OoxXml.TryParsePercentage(null, out _), "A missing percentage has no value.");
        TestAssert.True(!OoxXml.TryParsePercentage("2147483648", out _), "The Int32 caller contract is retained.");
        TestAssert.True(!OoxXml.TryParsePercentage("2147483.648%", out _), "Literal units retain the same numeric boundary.");
        TestAssert.Throws<OverflowException>(() => OoxXml.ParsePercentage("2147483648"));
        TestAssert.Equal(int.MinValue / 100000d, OoxXml.ParsePercentage(int.MinValue.ToString(CultureInfo.InvariantCulture)));
        TestAssert.Equal(int.MaxValue / 100000d, OoxXml.ParsePercentage(int.MaxValue.ToString(CultureInfo.InvariantCulture)));
    }

    public static void OoxPercentageFallbackKeepsColorIntegerRangeAndInvariantCulture()
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            foreach (string source in new[] { "12500", "12.5%" })
                TestAssert.Equal(.125d, OoxXml.ReadOptionalPercentage(new XElement("color", new XAttribute("val", source)), "val", .75d));
            foreach (string source in new[] { "bad", "12,5%", "NaN%", "Infinity%" })
                TestAssert.Equal(.75d, OoxXml.ReadOptionalPercentage(new XElement("color", new XAttribute("val", source)), "val", .75d));
            TestAssert.Equal(.75d, OoxXml.ReadOptionalPercentage(null, "val", .75d));
            TestAssert.Equal(.75d, OoxXml.ReadOptionalPercentage(new XElement("color"), "val", .75d));
            foreach (long source in new[] { long.MinValue, long.MaxValue })
                TestAssert.Equal(source / 100000d, OoxXml.ReadOptionalPercentage(
                    new XElement("color", new XAttribute("val", source.ToString(CultureInfo.InvariantCulture))), "val", .75d));
        }
        finally { CultureInfo.CurrentCulture = original; }
    }
}
