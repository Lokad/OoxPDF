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
    DocxRevisionInfo? Revision)
{
    public double? ParentFontSize { get; init; }
    public bool Cramped { get; init; }
}

internal abstract record DocxMathNode
{
    // Preserve authored math properties until their layout semantics are supported.
    public IReadOnlyList<string> PropertiesXml { get; init; } = [];
}

internal sealed record DocxMathArgumentProperties(int? Size, bool HasUnsupportedProperties);
internal sealed record DocxMathRow(IReadOnlyList<DocxMathNode> Children) : DocxMathNode
{
    public DocxMathArgumentProperties? ArgumentProperties { get; init; }
    public bool HasKnownArgumentProperties => ArgumentProperties is { HasUnsupportedProperties: false };
    public bool HasSupportedProperties => PropertiesXml.Count == 0 ||
        ArgumentProperties is { HasUnsupportedProperties: false, Size: null or 0 };
}
internal sealed record DocxMathRun(DocxTextRun Text, string? MathStyleValue, string? ScriptValue) : DocxMathNode
{
    public bool? NormalTextValue { get; init; }
    public bool NormalText => NormalTextValue == true;
    public bool HasUnsupportedProperties { get; init; }
    public string? ExplicitFontFamily { get; init; }
}
internal sealed record DocxMathFraction(DocxMathRow Numerator, DocxMathRow Denominator, string? TypeValue) : DocxMathNode
{
    public bool HasUnsupportedProperties { get; init; }
}
internal sealed record DocxMathScripts(DocxMathRow Base, DocxMathRow? Subscript, DocxMathRow? Superscript) : DocxMathNode;
internal sealed record DocxMathRadical(DocxMathRow Radicand, DocxMathRow Degree, bool? DegreeHidden) : DocxMathNode
{
    public bool HasUnsupportedProperties { get; init; }
    public DocxTextRun? ControlText { get; init; }
    public string? ControlFontFamily { get; init; }
}
internal sealed record DocxMathDelimiter(IReadOnlyList<DocxMathRow> Arguments) : DocxMathNode
{
    public string? BeginCharacterValue { get; init; }
    public string? EndCharacterValue { get; init; }
    public string? SeparatorCharacterValue { get; init; }
    public string BeginCharacter => BeginCharacterValue ?? "(";
    public string EndCharacter => EndCharacterValue ?? ")";
    public string SeparatorCharacter => SeparatorCharacterValue ?? "|";
    public bool? GrowValue { get; init; }
    public string? ShapeValue { get; init; }
    public bool HasUnsupportedProperties { get; init; }
    public DocxTextRun? ControlText { get; init; }
    public string? ControlFontFamily { get; init; }
}
internal sealed record DocxMathChild(string ElementName, string NamespaceName, DocxMathNode Content);
internal sealed record DocxMathUnsupported(string ElementName, string NamespaceName, string? SourceXml,
    IReadOnlyList<DocxMathChild> Children) : DocxMathNode;
