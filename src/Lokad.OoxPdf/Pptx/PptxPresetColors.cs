namespace Lokad.OoxPdf.Pptx;

internal static class PptxPresetColors
{
    private static readonly IReadOnlyDictionary<string, RgbColor> Colors = new Dictionary<string, RgbColor>(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = new(0x00, 0x00, 0x00),
        ["white"] = new(0xFF, 0xFF, 0xFF),
        ["red"] = new(0xFF, 0x00, 0x00),
        ["green"] = new(0x00, 0x80, 0x00),
        ["blue"] = new(0x00, 0x00, 0xFF),
        ["yellow"] = new(0xFF, 0xFF, 0x00),
        ["cyan"] = new(0x00, 0xFF, 0xFF),
        ["magenta"] = new(0xFF, 0x00, 0xFF),
        ["orange"] = new(0xFF, 0xA5, 0x00),
        ["purple"] = new(0x80, 0x00, 0x80),
        ["gray"] = new(0x80, 0x80, 0x80),
        ["grey"] = new(0x80, 0x80, 0x80),
        ["lime"] = new(0x00, 0xFF, 0x00),
        ["navy"] = new(0x00, 0x00, 0x80),
        ["teal"] = new(0x00, 0x80, 0x80),
        ["maroon"] = new(0x80, 0x00, 0x00),
        ["olive"] = new(0x80, 0x80, 0x00),
        ["silver"] = new(0xC0, 0xC0, 0xC0),
        ["aqua"] = new(0x00, 0xFF, 0xFF),
        ["fuchsia"] = new(0xFF, 0x00, 0xFF),
        ["darkBlue"] = new(0x00, 0x00, 0x8B),
        ["darkGreen"] = new(0x00, 0x64, 0x00),
        ["darkRed"] = new(0x8B, 0x00, 0x00),
        ["gold"] = new(0xFF, 0xD7, 0x00),
        ["cornflowerBlue"] = new(0x64, 0x95, 0xED)
    };

    public static bool TryResolve(string? name, out RgbColor color)
    {
        if (name is not null && Colors.TryGetValue(name, out color))
        {
            return true;
        }

        color = default;
        return false;
    }
}
