using System.Text;

namespace Lokad.OoxPdf.Ooxml;

/// <summary>
/// Shared numeral primitives for OOXML numbering labels.
/// </summary>
internal static class OoxNumbering
{
    /// <summary>
    /// Alphabetic label (A to Z, then AA onward) for positive values.
    /// </summary>
    internal static string ToAlphabetic(int value, bool upper)
    {
        var builder = new StringBuilder();
        int current = value;
        while (current > 0)
        {
            current--;
            char letter = (char)((upper ? 'A' : 'a') + current % 26);
            builder.Insert(0, letter);
            current /= 26;
        }

        return builder.ToString();
    }

    /// <summary>
    /// Upper-case Roman label for positive values; callers keep range and case contracts.
    /// </summary>
    internal static string ToRomanUpper(int value)
    {
        ReadOnlySpan<(int Value, string Text)> numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];
        var builder = new StringBuilder();
        int current = value;
        foreach ((int numeralValue, string numeralText) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(numeralText);
                current -= numeralValue;
            }
        }

        return builder.ToString();
    }
}
