namespace Lokad.OoxPdf.Docx;

// Boundaries refer to the existing prose indices. Mathematics does not become
// a prose run or change ownership of fields, notes, bookmarks or pictures.
internal sealed record DocxEquation(
    DocxMathRow Content,
    int TextRunIndex,
    int SourceRunIndex,
    bool Display,
    string? AlignmentValue,
    string? PropertiesXml,
    DocxRevisionInfo? Revision);

internal abstract record DocxMathNode
{
    // Preserve authored math properties until their layout semantics are supported.
    public IReadOnlyList<string> PropertiesXml { get; init; } = [];
}

internal sealed record DocxMathRow(IReadOnlyList<DocxMathNode> Children) : DocxMathNode;
internal sealed record DocxMathRun(DocxTextRun Text, string? MathStyleValue, string? ScriptValue) : DocxMathNode
{
    public bool HasUnsupportedProperties { get; init; }
    public string? ExplicitFontFamily { get; init; }
}
internal sealed record DocxMathFraction(DocxMathRow Numerator, DocxMathRow Denominator, string? TypeValue) : DocxMathNode
{
    public bool HasUnsupportedProperties { get; init; }
}
internal sealed record DocxMathScripts(DocxMathRow Base, DocxMathRow? Subscript, DocxMathRow? Superscript) : DocxMathNode;
internal sealed record DocxMathChild(string ElementName, string NamespaceName, DocxMathNode Content);
internal sealed record DocxMathUnsupported(string ElementName, string NamespaceName, string? SourceXml,
    IReadOnlyList<DocxMathChild> Children) : DocxMathNode;
