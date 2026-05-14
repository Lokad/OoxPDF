namespace Lokad.OoxPdf.Ooxml;

internal sealed record OoxRelationship(
    string Id,
    string Type,
    string Target,
    string? TargetMode,
    string? ResolvedTarget)
{
    public bool IsExternal => TargetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true;
}
