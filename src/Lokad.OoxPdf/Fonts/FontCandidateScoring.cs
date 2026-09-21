namespace Lokad.OoxPdf.Fonts;

/// <summary>
/// Shared font-candidate ordering mechanics (D02). Every resolver ranks style
/// fidelity first (italic match, then weight distance); population-specific
/// tie-breakers (family spelling, source identity, face index) stay local to
/// each resolver with explicit comparers so Word/Pack/Windows policies cannot
/// drift silently.
/// </summary>
internal static class FontCandidateScoring
{
    public static int CompareStyleWeight(
        bool leftItalic,
        int leftWeightClass,
        bool rightItalic,
        int rightWeightClass,
        bool requestItalic,
        int targetWeight)
    {
        int result = (leftItalic == requestItalic ? 0 : 1000).CompareTo(rightItalic == requestItalic ? 0 : 1000);
        if (result != 0)
        {
            return result;
        }

        return Math.Abs(leftWeightClass - targetWeight).CompareTo(Math.Abs(rightWeightClass - targetWeight));
    }
}
