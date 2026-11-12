using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathParser
{
    // The package XML limits still apply. Use an explicit traversal stack so
    // accepted XML nesting does not also consume the managed call stack.
    public static DocxMathRow Read(
        XElement equation,
        Func<XElement, string, DocxTextRun> readRun,
        CancellationToken cancellationToken, Func<XElement, bool>? includeRevision = null,
        bool preserveDeletedStructure = false, Action<DocxMathBookmark>? readBookmark = null,
        Action<XElement>? readCommentBoundary = null, Func<XElement, DocxHyperlinkSpan>? readHyperlink = null)
    {
        var pending = new Stack<(XElement Element, bool Expanded, bool InsideRevision, bool IncludeAnnotations, DocxHyperlinkSpan? Hyperlink)>();
        var nodes = new Dictionary<XElement, DocxMathNode>();
        pending.Push((equation, false, false, true, null));
        while (pending.TryPop(out var item))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = item.Element;
            string name = element.Name.LocalName;
            if (IsCommentBoundary(element))
            {
                if (item.IncludeAnnotations) readCommentBoundary?.Invoke(element);
                continue;
            }
            bool revisionSource = !item.InsideRevision && includeRevision is not null && element.Name.Namespace == WordprocessingNamespace &&
                name is "ins" or "del";
            XElement[]? revisionChildren = revisionSource ? ReadRevisionChildren(element, cancellationToken).ToArray() : null;
            bool revisionRow = revisionChildren is { Length: > 0 } && revisionChildren.All(IsRevisionContent);
            bool included = !revisionRow || includeRevision!(element);
            bool retainDeletedFrame = revisionRow && preserveDeletedStructure && name == "del" && !included &&
                revisionChildren!.Any(child => IsControl(child.Name.LocalName));
            if (revisionRow && !item.Expanded && !included && !retainDeletedFrame)
            {
                nodes.Add(element, new DocxMathRow([]));
                continue;
            }
            bool sourceContainer = IsSourceContainer(element);
            if (!revisionRow && !sourceContainer && (element.Name.Namespace != MathNamespace || !IsContainer(name)))
            {
                nodes.Add(element, ReadLeaf(element, readRun) with { Hyperlink = item.IncludeAnnotations ? item.Hyperlink : null });
                continue;
            }

            if (!item.Expanded)
            {
                DocxHyperlinkSpan? hyperlink = item.IncludeAnnotations && included ? item.Hyperlink : null;
                if (item.IncludeAnnotations && included && element.Name == WordprocessingNamespace + "hyperlink")
                    hyperlink = readHyperlink?.Invoke(element);
                pending.Push((element, true, item.InsideRevision, item.IncludeAnnotations && included, hyperlink));
                // Comment ranges need source order, including boundaries in
                // different nested rows. They are events, not math operands.
                foreach (XElement child in TraversalChildren(element).Reverse())
                {
                    pending.Push((child, false, item.InsideRevision || revisionRow,
                        item.IncludeAnnotations && included, hyperlink));
                }
                continue;
            }

            DocxMathNode node;
            if (revisionRow || sourceContainer)
                node = new DocxMathRow(ContentChildren(element).Select(child => nodes[child]).ToArray());
            else if (IsRow(name) && (name == "oMath" || !IsSourceContainer(element.Parent)) &&
                (name != "deg" || element.Parent?.Name == MathNamespace + "rad") &&
                (name != "fName" || element.Parent?.Name == MathNamespace + "func"))
            {
                node = new DocxMathRow(ContentChildren(element).Select(child => nodes[child]).ToArray())
                { ArgumentProperties = ReadArgumentProperties(element) };
            }
            else if (name == "mr" && element.Parent?.Name == MathNamespace + "m" && ContentChildren(element).Any() &&
                ContentChildren(element).All(child => child.Name == MathNamespace + "e" && nodes[child] is DocxMathRow))
            {
                node = new DocxMathMatrixRow(ContentChildren(element).Select(child => (DocxMathRow)nodes[child]).ToArray())
                { HasUnsupportedProperties = element.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                    !HasOnlyWhitespaceText(element) || element.Elements().Any(IsProperty) };
            }
            else if (name == "m" && ContentChildren(element).Any() &&
                ContentChildren(element).All(child => child.Name == MathNamespace + "mr" && nodes[child] is DocxMathMatrixRow))
            {
                node = ReadMatrix(element, ContentChildren(element).Select(child => (DocxMathMatrixRow)nodes[child]).ToArray(),
                    readRun, cancellationToken);
            }
            else if (name == "f" && Role(element, "num") is { } numerator && Role(element, "den") is { } denominator &&
                HasOnlyRoles(element, "num", "den"))
            {
                node = ReadFraction(element, numerator, denominator, readRun, cancellationToken);
            }
            else if (name is "sSub" or "sSup" or "sSubSup" && Role(element, "e") is { } basis &&
                (name == "sSup" || Role(element, "sub") is not null) &&
                (name == "sSub" || Role(element, "sup") is not null) &&
                HasOnlyRoles(element, name == "sSub" ? ["e", "sub"] : name == "sSup" ? ["e", "sup"] : ["e", "sub", "sup"]))
            {
                node = ReadScripts(element, basis, Role(element, "sub"), Role(element, "sup"), readRun, cancellationToken);
            }
            else if (name == "func" && Role(element, "fName") is { } functionName && Role(element, "e") is { } argument &&
                HasOnlyRoles(element, "fName", "e"))
            {
                node = ReadFunction(element, functionName, argument, readRun, cancellationToken);
            }
            else if (name == "nary" && Role(element, "e") is { } body && Role(element, "sub") is { } lower &&
                Role(element, "sup") is { } upper && HasOnlyRoles(element, "sub", "sup", "e"))
            {
                node = ReadNary(element, body, lower, upper, readRun, cancellationToken);
            }
            else if (name == "d" && ContentChildren(element).Any() &&
                ContentChildren(element).All(child => child.Name == MathNamespace + "e" && nodes[child] is DocxMathRow))
            {
                node = ReadDelimiter(element, element.Elements(MathNamespace + "e").Select(child => (DocxMathRow)nodes[child]).ToArray(), readRun, cancellationToken);
            }
            else if (name == "rad" && Role(element, "e") is { } radicand && Role(element, "deg") is { } degree &&
                HasOnlyRoles(element, "deg", "e"))
            {
                XElement[] properties = element.Elements().Where(IsProperty).ToArray();
                XElement[] visibility = element.Element(MathNamespace + "radPr")?.Elements(MathNamespace + "degHide").ToArray() ?? [];
                XElement[] controls = element.Element(MathNamespace + "radPr")?.Elements(MathNamespace + "ctrlPr").ToArray() ?? [];
                XElement? control = controls.FirstOrDefault();
                DocxTextRun controlText = readRun(control ?? element, string.Empty);
                string? value = (string?)visibility.FirstOrDefault()?.Attribute(MathNamespace + "val");
                bool valid = properties.Length <= 1 && properties.All(p => p.Name == MathNamespace + "radPr" &&
                    p.Attributes().All(a => a.IsNamespaceDeclaration) &&
                    string.IsNullOrWhiteSpace(string.Concat(p.Nodes().OfType<XText>().Select(t => t.Value))) &&
                    p.Elements().All(child => child.Name == MathNamespace + "degHide" || child.Name == MathNamespace + "ctrlPr")) &&
                    controls.Length <= 1 && controls.All(HasKnownControl) &&
                    visibility.Length <= 1 && visibility.All(v => !v.HasElements && string.IsNullOrWhiteSpace(v.Value) &&
                        v.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == MathNamespace + "val")) &&
                    (value is null || Lokad.OoxPdf.Ooxml.OoxBoolean.IsTrue(value) || Lokad.OoxPdf.Ooxml.OoxBoolean.IsOff(value));
                node = new DocxMathRadical(radicand, degree, visibility.Length == 0 ? null :
                    Lokad.OoxPdf.Ooxml.OoxBoolean.ParseElement(visibility[0], false, valueAttributeName: MathNamespace + "val"))
                    { HasUnsupportedProperties = !valid, ControlText = controlText,
                        ControlFontFamily = ControlFontFamily(control, controlText) };
            }
            else
            {
                // Do not serialize already traversed subtrees again: malformed
                // nested constructs must retain linear work and storage.
                node = new DocxMathUnsupported(name, element.Name.NamespaceName, null,
                    ContentChildren(element).Select(child => new DocxMathChild(
                        child.Name.LocalName, child.Name.NamespaceName, nodes[child])).ToArray());
            }
            if (retainDeletedFrame)
            {
                var retained = DocxMathLayout.RemoveDeletedLiteralMathText(
                    new(new([node]), 0, 0, false, null, null, null), cancellationToken);
                node = retained is not null ? retained.Children.Count == 1 ? retained.Children[0] : retained :
                    new DocxMathUnsupported(name, element.Name.NamespaceName, null,
                        ContentChildren(element).Select(child => new DocxMathChild(
                            child.Name.LocalName, child.Name.NamespaceName, nodes[child])).ToArray());
            }
            if (node is DocxMathRow sourceRow && item.IncludeAnnotations)
            {
                List<DocxMathBookmarkBoundary>? bookmarks = null;
                int childIndex = 0;
                foreach (var child in SourceChildren(element))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (IsBookmarkMarker(child))
                    {
                        if (child.Name.LocalName == "bookmarkStart")
                        {
                            var bookmark = new DocxMathBookmark((string?)child.Attribute(WordprocessingNamespace + "id"),
                                (string?)child.Attribute(WordprocessingNamespace + "name"));
                            (bookmarks ??= []).Add(new(bookmark, childIndex));
                            readBookmark?.Invoke(bookmark);
                        }
                    }
                    else if (!IsProperty(child) && !IsNonPrintingMarker(child) && !IsCommentBoundary(child)) childIndex++;
                }
                if (bookmarks is not null) node = sourceRow with { Bookmarks = bookmarks };
            }
            node = node with { Hyperlink = item.IncludeAnnotations && included ? item.Hyperlink : null };
            nodes.Add(element, revisionRow ? node : node with { PropertiesXml = ReadProperties(element) });
        }

        return nodes[equation] as DocxMathRow ?? new DocxMathRow([nodes[equation]]);

        DocxMathRow? Role(XElement element, string name)
        {
            XElement[] matches = element.Elements(MathNamespace + name).ToArray();
            return matches.Length == 1 ? nodes[matches[0]] as DocxMathRow : null;
        }
    }

    private static bool IsRow(string name) => name is "oMath" or "e" or "sub" or "sup" or "num" or "den" or "deg" or "fName";
    private static bool HasKnownControlProperty(XElement property) =>
        property.Name.Namespace == WordprocessingNamespace && !property.HasElements &&
        string.IsNullOrWhiteSpace(property.Value) && property.Name.LocalName is
            "rFonts" or "sz" or "szCs" or "color" or "i" or "iCs" or "b" or "bCs" or "lang" or "rStyle" or
            "u" or "strike" or "dstrike" or "vanish" or "caps" or "smallCaps" or "vertAlign" or "spacing" or "highlight" or "shd";
    // A single control has the same grammar inside an admitted revision wrapper.
    // Role rows and orphan matrix rows still require their actual source parent.
    private static bool IsControl(string name) => name is "f" or "sSub" or "sSup" or "sSubSup" or "rad" or "d" or "nary" or "func" or "m";
    private static bool IsRevisionContent(XElement element) => element.Name.Namespace == MathNamespace &&
        (IsControl(element.Name.LocalName) || element.Name.LocalName == "r" &&
            (element.Elements(MathNamespace + "t").Any(text => text.Value.Length != 0) ||
             IsReadableRun(element) && HasSupportedRunProperties(element)));
    // Revision admission and structural retention see the same stored content
    // through metadata containers; role nodes still need their actual parent.
    private static IEnumerable<XElement> ReadRevisionChildren(XElement source, CancellationToken token)
    {
        var pending = new Stack<XElement>(NonMarkerChildren(source).Reverse());
        while (pending.TryPop(out var child))
        {
            token.ThrowIfCancellationRequested();
            if (IsSourceContainer(child))
            {
                foreach (var item in NonMarkerChildren(child).Reverse()) pending.Push(item);
            }
            else yield return child;
        }
    }
    private static bool IsSourceContainer(XElement? element) =>
        element?.Name == WordprocessingNamespace + "hyperlink" ||
        element?.Name == WordprocessingNamespace + "customXml" || element?.Name == WordprocessingNamespace + "sdt" ||
        element?.Name == WordprocessingNamespace + "sdtContent" && element.Parent?.Name == WordprocessingNamespace + "sdt";
    // Match paragraph reading: control/custom-XML properties describe metadata;
    // only stored child content supplies mathematical operands.
    private static IEnumerable<XElement> SourceChildren(XElement element) =>
        element.Name == WordprocessingNamespace + "sdt"
            ? element.Elements(WordprocessingNamespace + "sdtContent")
            : element.Name == WordprocessingNamespace + "customXml"
                ? element.Elements().Where(child => child.Name != WordprocessingNamespace + "customXmlPr")
                : element.Elements();
    private static bool IsContainer(string name) => IsRow(name) || IsControl(name) || name == "mr";
    private static bool IsProperty(XElement element) => element.Name.Namespace == MathNamespace && element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal);
    // Editing boundaries and stored pagination positions have no operand or ink.
    // Keep content-bearing, unknown and misplaced payloads on the unsupported path.
    private static bool IsNonPrintingMarker(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        (element.Name.LocalName is "proofErr" or "permStart" or "permEnd" ||
            element.Name.LocalName == "lastRenderedPageBreak" && element.Parent?.Name == MathNamespace + "r" &&
            element.Attributes().All(attribute => attribute.IsNamespaceDeclaration));
    private static bool IsBookmarkMarker(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "bookmarkStart" or "bookmarkEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);
    private static bool IsCommentBoundary(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "commentRangeStart" or "commentRangeEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);
    private static IEnumerable<XElement> TraversalChildren(XElement element) =>
        SourceChildren(element).Where(child => !IsProperty(child) && !IsNonPrintingMarker(child) && !IsBookmarkMarker(child));
    private static IEnumerable<XElement> NonMarkerChildren(XElement element) =>
        SourceChildren(element).Where(child => !IsNonPrintingMarker(child) && !IsBookmarkMarker(child) && !IsCommentBoundary(child));
    private static IEnumerable<XElement> ContentChildren(XElement element) =>
        NonMarkerChildren(element).Where(child => !IsProperty(child));
    private static bool HasOnlyRoles(XElement element, params string[] roles)
    {
        XElement[] children = ContentChildren(element).ToArray();
        return children.Length == roles.Length && children.All(child => child.Name.Namespace == MathNamespace && roles.Contains(child.Name.LocalName));
    }
    private static string? PropertyValue(XElement element, string property, string value) =>
        (string?)element.Element(MathNamespace + property)?.Element(MathNamespace + value)?.Attribute(MathNamespace + "val");
    private static string[] ReadProperties(XElement element) => SourceChildren(element).Where(IsProperty)
        .Select(child => child.ToString(SaveOptions.DisableFormatting)).ToArray();

    private static DocxMathArgumentProperties? ReadArgumentProperties(XElement element)
    {
        XElement[] properties = element.Elements().Where(IsProperty).ToArray();
        if (properties.Length == 0) { return null; }
        if (properties.Length != 1 || properties[0].Name != MathNamespace + "argPr")
            return new(null, true);
        XElement argument = properties[0];
        XElement[] children = argument.Elements().ToArray();
        if (argument.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            !string.IsNullOrWhiteSpace(string.Concat(argument.Nodes().OfType<XText>().Select(text => text.Value))))
            return new(null, true);
        if (children.Length == 0) { return new(null, false); }
        if (children.Length != 1 || children[0].Name != MathNamespace + "argSz") { return new(null, true); }
        XElement size = children[0];
        if (size.HasElements || !string.IsNullOrWhiteSpace(size.Value) ||
            size.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != MathNamespace + "val"))
            return new(null, true);
        string? source = (string?)size.Attribute(MathNamespace + "val");
        int value = 0;
        if (source is not null && (!int.TryParse(source,
            System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value) || value is < -2 or > 2))
            return new(null, true);
        return new(value, false);
    }

    // Revision admission and leaf parsing share the same zero-content run shape.
    // Visible unknown payloads retain their former deferred fallback; an empty
    // unknown payload cannot become a supported revision owner.
    private static bool IsReadableRun(XElement element) =>
        element.Name == MathNamespace + "r" && element.Elements().All(child =>
            child.Name == MathNamespace + "t" || child.Name == MathNamespace + "rPr" || child.Name == WordprocessingNamespace + "rPr" ||
            IsNonPrintingMarker(child));

    private static bool HasSupportedRunProperties(XElement element) =>
        element.Element(MathNamespace + "rPr")?.Elements().Any(property =>
            property.Name != MathNamespace + "sty" && property.Name != MathNamespace + "scr" &&
            property.Name != MathNamespace + "nor") != true && HasSupportedNormalTextProperty(element);

    private static DocxMathNode ReadLeaf(XElement element, Func<XElement, string, DocxTextRun> readRun)
    {
        if (IsReadableRun(element))
        {
            return new DocxMathRun(
                readRun(element, string.Concat(element.Elements(MathNamespace + "t").Select(text => text.Value))),
                PropertyValue(element, "rPr", "sty"), PropertyValue(element, "rPr", "scr"))
            {
                PropertiesXml = ReadProperties(element),
                NormalTextValue = element.Element(MathNamespace + "rPr")?.Element(MathNamespace + "nor") is { } authoredNormal
                    ? Lokad.OoxPdf.Ooxml.OoxBoolean.ParseElement(authoredNormal, false, valueAttributeName: MathNamespace + "val") : null,
                ExplicitFontFamily = (string?)element.Element(WordprocessingNamespace + "rPr")?.Element(WordprocessingNamespace + "rFonts")?.Attribute(WordprocessingNamespace + "ascii")
                    ?? (string?)element.Element(WordprocessingNamespace + "rPr")?.Element(WordprocessingNamespace + "rFonts")?.Attribute(WordprocessingNamespace + "hAnsi"),
                HasUnsupportedProperties = !HasSupportedRunProperties(element)
            };
        }
        return Unsupported(element);
    }

    private static bool HasSupportedNormalTextProperty(XElement element)
    {
        XElement[] properties = element.Element(MathNamespace + "rPr")?.Elements(MathNamespace + "nor").ToArray() ?? [];
        if (properties.Length == 0) { return true; }
        if (properties.Length != 1 || properties[0].HasElements || !string.IsNullOrWhiteSpace(properties[0].Value) ||
            properties[0].Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != MathNamespace + "val"))
            return false;
        string? value = (string?)properties[0].Attribute(MathNamespace + "val");
        return value is null || Lokad.OoxPdf.Ooxml.OoxBoolean.IsTrue(value) || Lokad.OoxPdf.Ooxml.OoxBoolean.IsOff(value);
    }

    private static DocxMathUnsupported Unsupported(XElement element) =>
        new(element.Name.LocalName, element.Name.NamespaceName, element.ToString(SaveOptions.DisableFormatting), []);
}
