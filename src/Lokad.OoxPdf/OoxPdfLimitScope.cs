namespace Lokad.OoxPdf;

/// <summary>Identifies the ownership scope of a resource refusal.</summary>
public enum OoxPdfLimitScope
{
    /// <summary>Storage and format bounds belonging to the current conversion.</summary>
    Conversion,
    /// <summary>A shared resource whose ownership can outlive one conversion.</summary>
    SharedResource
}
