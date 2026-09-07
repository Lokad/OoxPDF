namespace Lokad.OoxPdf.Docx;

internal static class DocxDefaults
{
    public const double FontSizePoints = 11d;

    // Word sizes runs with no resolvable direct, style, table, or default size at flat 12pt
    // (bodysize probes 2026-09-07: styles-less and styles-without-size-or-defaults emit 12pt bodies
    // at unsized runs while sized runs keep direct size).
    public const double UnstyledRunFontSizePoints = 12d;
}
