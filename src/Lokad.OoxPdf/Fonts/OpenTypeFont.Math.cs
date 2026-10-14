namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    public bool TryGetMathConstants(out OpenTypeMathConstants constants)
    {
        constants = default;
        return tables.TryGetValue("MATH", out TableRecord table) &&
            OpenTypeMathConstants.TryRead(bytes.AsMemory(table.Offset, table.Length), out constants);
    }
}
