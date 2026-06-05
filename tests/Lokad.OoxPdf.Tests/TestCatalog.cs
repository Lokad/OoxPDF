using System.Reflection;

namespace Lokad.OoxPdf.Tests;

internal sealed record TestCase(string Group, Action Action);

internal static class TestCatalog
{
    public static IReadOnlyList<TestCase> All { get; } =
    [
        .. FromType("api", typeof(PublicApiTests)),
        .. FromType("cli", typeof(CliTests)),
        .. FromType("ooxml", typeof(OoxmlTests)),
        .. FromType("pdf", typeof(PdfWriterTests)),
        .. FromType(ClassifyPptx, typeof(PptxTests)),
        .. FromType(ClassifyDocx, typeof(DocxTests)),
        .. FromType("imaging", typeof(ImagingTests)),
        .. FromType("fonts", typeof(FontTests))
    ];

    private static IReadOnlyList<TestCase> FromType(string group, Type type)
    {
        return FromType(_ => group, type);
    }

    private static IReadOnlyList<TestCase> FromType(Func<string, string> classify, Type type)
    {
        return type
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.ReturnType == typeof(void) && method.GetParameters().Length == 0)
            .OrderBy(method => method.MetadataToken)
            .Select(method => new TestCase(classify(method.Name), (Action)Delegate.CreateDelegate(typeof(Action), method)))
            .ToArray();
    }

    private static string ClassifyPptx(string name)
    {
        if (name.Contains("Grouped", StringComparison.Ordinal) ||
            name.Contains("Order", StringComparison.Ordinal) ||
            name.Contains("Sibling", StringComparison.Ordinal) ||
            name.Contains("Above", StringComparison.Ordinal))
        {
            return "pptx-composition";
        }

        if (name.Contains("Chart", StringComparison.Ordinal))
        {
            return "pptx-charts";
        }

        if (name.Contains("Scene", StringComparison.Ordinal) ||
            name.Contains("Theme", StringComparison.Ordinal) ||
            name.Contains("Layout", StringComparison.Ordinal) ||
            name.Contains("Master", StringComparison.Ordinal) ||
            name.Contains("Placeholder", StringComparison.Ordinal))
        {
            return "pptx-model";
        }

        if (name.Contains("Table", StringComparison.Ordinal))
        {
            return "pptx-tables";
        }

        if (name.Contains("Text", StringComparison.Ordinal) ||
            name.Contains("Font", StringComparison.Ordinal) ||
            name.Contains("Bullet", StringComparison.Ordinal) ||
            name.Contains("Paragraph", StringComparison.Ordinal) ||
            name.Contains("Tab", StringComparison.Ordinal))
        {
            return "pptx-typography";
        }

        if (name.Contains("Picture", StringComparison.Ordinal) ||
            name.Contains("Image", StringComparison.Ordinal) ||
            name.Contains("Png", StringComparison.Ordinal) ||
            name.Contains("Bmp", StringComparison.Ordinal) ||
            name.Contains("Crop", StringComparison.Ordinal))
        {
            return "pptx-images";
        }

        if (name.Contains("Shape", StringComparison.Ordinal) ||
            name.Contains("Geometry", StringComparison.Ordinal) ||
            name.Contains("Connector", StringComparison.Ordinal) ||
            name.Contains("Arc", StringComparison.Ordinal) ||
            name.Contains("Arrow", StringComparison.Ordinal))
        {
            return "pptx-shapes";
        }

        return "pptx-core";
    }

    private static string ClassifyDocx(string name)
    {
        if (name.Contains("Table", StringComparison.Ordinal))
        {
            return "docx-tables";
        }

        if (name.Contains("Header", StringComparison.Ordinal) ||
            name.Contains("Footer", StringComparison.Ordinal) ||
            name.Contains("Page", StringComparison.Ordinal))
        {
            return "docx-page";
        }

        if (name.Contains("Numbering", StringComparison.Ordinal))
        {
            return "docx-numbering";
        }

        if (name.Contains("Image", StringComparison.Ordinal) ||
            name.Contains("Png", StringComparison.Ordinal))
        {
            return "docx-images";
        }

        if (name.Contains("Paragraph", StringComparison.Ordinal) ||
            name.Contains("LineHeight", StringComparison.Ordinal) ||
            name.Contains("Style", StringComparison.Ordinal))
        {
            return "docx-text";
        }

        return "docx-core";
    }
}
