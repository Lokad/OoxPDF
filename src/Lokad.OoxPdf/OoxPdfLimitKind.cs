namespace Lokad.OoxPdf;

/// <summary>Identifies why a conversion cannot continue within its resource contract.</summary>
public enum OoxPdfLimitKind
{
    /// <summary>An input, format or output bound was exceeded.</summary>
    DocumentBound,
    /// <summary>The host refused admission to shared memory capacity.</summary>
    MemoryAdmission
}
