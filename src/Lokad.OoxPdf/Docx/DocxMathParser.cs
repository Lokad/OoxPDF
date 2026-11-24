using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
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
        Action<XElement>? readCommentBoundary = null, Func<XElement, DocxHyperlinkSpan>? readHyperlink = null,
        XElement? authoredStory = null, bool equationHasDeletionOwner = false,
        (XElement Run, XElement? FirstBreakInRun)? trailingBreaks = null)
    {
        var pending = new Stack<(XElement Element, bool Expanded, XName? RevisionOwner, bool IncludeAnnotations, DocxHyperlinkSpan? Hyperlink)>();
        var nodes = new Dictionary<XElement, DocxMathNode>();
        IReadOnlySet<(XName Name, string Id)>? moveOwners = null;
        XName deletion = WordprocessingNamespace + "del";
        pending.Push((equation, false, null, true, null));
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
            bool revisionSource = item.RevisionOwner is null && includeRevision is not null &&
                DocxSourceXml.IsRevisionContainer(element) &&
                (name is not ("moveFrom" or "moveTo") ||
                    (string?)element.Attribute(WordprocessingNamespace + "id") is { } id &&
                    (moveOwners ??= DocxSourceXml.ReadPairedMoveOwnerIds(equation, authoredStory, cancellationToken))
                        .Contains((element.Name, id)));
            XElement[]? revisionChildren = revisionSource ? ReadRevisionChildren(element, cancellationToken).ToArray() : null;
            bool deleting = element.Name == deletion;
            bool revisionRow = revisionChildren is not null && IsReadableRevision(revisionChildren, deleting);
            bool included = !revisionRow || includeRevision!(element);
            bool retainDeletedFrame = revisionRow && preserveDeletedStructure && deleting && !included &&
                revisionChildren!.Any(child => IsControl(child.Name.LocalName));
            if (revisionRow && !item.Expanded && !included && !retainDeletedFrame)
            {
                nodes.Add(element, new DocxMathRow([]));
                continue;
            }
            bool sourceContainer = IsSourceContainer(element);
            if (!revisionRow && !sourceContainer && (element.Name.Namespace != MathNamespace || !IsContainer(name)))
            {
                bool allowsDeletedText = item.RevisionOwner is null ? equationHasDeletionOwner : item.RevisionOwner == deletion;
                XElement? firstBreakInRun = element == trailingBreaks?.Run ? trailingBreaks?.FirstBreakInRun : null;
                nodes.Add(element, ReadLeaf(element, readRun, allowsDeletedText, firstBreakInRun) with
                    { Hyperlink = item.IncludeAnnotations ? item.Hyperlink : null });
                continue;
            }

            if (!item.Expanded)
            {
                DocxHyperlinkSpan? hyperlink = item.IncludeAnnotations && included ? item.Hyperlink : null;
                if (item.IncludeAnnotations && included && element.Name == WordprocessingNamespace + "hyperlink")
                    hyperlink = readHyperlink?.Invoke(element);
                pending.Push((element, true, item.RevisionOwner, item.IncludeAnnotations && included, hyperlink));
                XName? revisionOwner = revisionRow ? element.Name : item.RevisionOwner;
                // Comment ranges need source order, including boundaries in
                // different nested rows. They are events, not math operands.
                foreach (XElement child in WithoutTrailingBreaks(element, TraversalChildren(element)).Reverse())
                {
                    pending.Push((child, false, revisionOwner,
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
                node = new DocxMathRow(WithoutTrailingBreaks(element, ContentChildren(element)).Select(child => nodes[child]).ToArray())
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
                node = new DocxMathUnsupported(name, element.Name.NamespaceName,
                    ContentChildren(element).Select(child => new DocxMathChild(
                        child.Name.LocalName, child.Name.NamespaceName, nodes[child])).ToArray());
            }
            if (retainDeletedFrame)
            {
                var retained = DocxMathLayout.RemoveDeletedLiteralMathText(
                    new(new([node]), 0, 0, false, null, null), cancellationToken);
                node = retained is not null ? retained.Children.Count == 1 ? retained.Children[0] : retained :
                    new DocxMathUnsupported(name, element.Name.NamespaceName,
                        ContentChildren(element).Select(child => new DocxMathChild(
                            child.Name.LocalName, child.Name.NamespaceName, nodes[child])).ToArray());
            }
            if (node is DocxMathRow sourceRow && item.IncludeAnnotations)
            {
                List<DocxMathBookmarkBoundary>? bookmarks = null;
                int childIndex = 0;
                foreach (var child in WithoutTrailingBreaks(element, DocxSourceXml.StoredChildren(element)))
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
                    else if (!IsProperty(child) && !IsNonPrintingSourceMarker(child) && !IsCommentBoundary(child)) childIndex++;
                }
                if (bookmarks is not null) node = sourceRow with { Bookmarks = bookmarks };
            }
            node = node with { Hyperlink = item.IncludeAnnotations && included ? item.Hyperlink : null };
            nodes.Add(element, revisionRow ? node : node with { PropertiesXml = ReadProperties(element) });
        }

        return nodes[equation] as DocxMathRow ?? new DocxMathRow([nodes[equation]]);

        IEnumerable<XElement> WithoutTrailingBreaks(XElement element, IEnumerable<XElement> children) =>
            element == equation && trailingBreaks is { } trailing ? BeforeBreakRun(children, trailing) : children;

        static IEnumerable<XElement> BeforeBreakRun(IEnumerable<XElement> children,
            (XElement Run, XElement? FirstBreakInRun) trailing)
        {
            foreach (XElement child in children)
            {
                if (child == trailing.Run)
                {
                    if (trailing.FirstBreakInRun is not null) yield return child;
                    yield break;
                }
                yield return child;
            }
        }

        DocxMathRow? Role(XElement element, string name)
        {
            XElement[] matches = element.Elements(MathNamespace + name).ToArray();
            return matches.Length == 1 ? nodes[matches[0]] as DocxMathRow : null;
        }
    }

    private static bool IsRow(string name) => name is "oMath" or "e" or "sub" or "sup" or "num" or "den" or "deg" or "fName";
    private static bool HasKnownControlProperty(XElement property)
    {
        if (property.Name == WordprocessingNamespace + "rPrChange")
        {
            // The nested properties describe the former appearance. The
            // current control paint comes from the outer run properties.
            return HasOnlyWhitespaceText(property) &&
                property.Element(WordprocessingNamespace + "rPr") is not null &&
                property.Elements().Take(2).Count() == 1;
        }

        return property.Name.Namespace == WordprocessingNamespace && !property.HasElements &&
            string.IsNullOrWhiteSpace(property.Value) && (property.Name.LocalName switch
            {
                "rFonts" or "sz" or "szCs" or "color" or "i" or "iCs" or "b" or "bCs" or "lang" or "rStyle" or
                    "u" or "strike" or "dstrike" or "vanish" or "caps" or "smallCaps" or "vertAlign" or "spacing" or "highlight" or "shd" => true,
                // Proofing and web-view flags preserve the control's paginated paint.
                "noProof" or "webHidden" => HasBooleanValue((string?)property.Attribute(WordprocessingNamespace + "val")) &&
                    property.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == WordprocessingNamespace + "val"),
                _ => false
            });
    }
    // A single control has the same grammar inside an admitted revision wrapper.
    // Role rows and orphan matrix rows still require their actual source parent.
    private static bool IsControl(string name) => name is "f" or "sSub" or "sSup" or "sSubSup" or "rad" or "d" or "nary" or "func" or "m";

    private static bool IsReadableRevision(XElement[] children, bool allowsDeletedText)
    {
        // A revision needs an actual operand or structural control.
        // Readable zero-content siblings do not invalidate that owner.
        bool hasContent = false;
        foreach (var child in children)
        {
            if (IsRevisionContent(child, allowsDeletedText))
                hasContent = true;
            else if (!IsReadableRun(child, allowsDeletedText) || !ReadRunProperties(child).Supported)
                return false;
        }
        return hasContent;
    }

    private static bool IsRevisionContent(XElement element, bool allowsDeletedText) => element.Name.Namespace == MathNamespace &&
        (IsControl(element.Name.LocalName) || element.Name.LocalName == "r" &&
            element.Elements().Any(text => ReadLiteralRunText(text, allowsDeletedText)?.Any(IsMathLiteralCharacter) == true));

    private static bool IsMathLiteralCharacter(char value) => !DocxTextBreakRules.IsOptionalBreakMarker(value);
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
        element is not null &&
        (DocxSourceXml.IsInlineMetadataContainer(element) ||
         element.Name == WordprocessingNamespace + "hyperlink" || element.Name == WordprocessingNamespace + "sdt" ||
         element.Name == WordprocessingNamespace + "sdtContent" && element.Parent?.Name == WordprocessingNamespace + "sdt");
    private static bool IsContainer(string name) => IsRow(name) || IsControl(name) || name == "mr";
    private static bool IsProperty(XElement element) => element.Name.Namespace == MathNamespace && element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal);
    internal static bool IsBookmarkMarker(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "bookmarkStart" or "bookmarkEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);
    internal static bool IsCommentBoundary(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "commentRangeStart" or "commentRangeEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);
    private static bool IsNonPrintingSourceMarker(XElement element) =>
        DocxSourceXml.IsNonPrintingMarker(element) ||
        DocxSourceXml.TryResolveRevisionRangeMarker(element, out _, out _) &&
        DocxSourceXml.IsNonPrintingRevisionRangeBoundary(element);

    private static IEnumerable<XElement> TraversalChildren(XElement element) =>
        DocxSourceXml.StoredChildren(element).Where(child => !IsProperty(child) && !IsNonPrintingSourceMarker(child) && !IsBookmarkMarker(child));
    private static IEnumerable<XElement> NonMarkerChildren(XElement element) =>
        DocxSourceXml.StoredChildren(element).Where(child => !IsNonPrintingSourceMarker(child) && !IsBookmarkMarker(child) && !IsCommentBoundary(child));
    private static IEnumerable<XElement> ContentChildren(XElement element) =>
        NonMarkerChildren(element).Where(child => !IsProperty(child));
    private static bool HasOnlyRoles(XElement element, params string[] roles)
    {
        XElement[] children = ContentChildren(element).ToArray();
        return children.Length == roles.Length && children.All(child => child.Name.Namespace == MathNamespace && roles.Contains(child.Name.LocalName));
    }
    private static string? PropertyValue(XElement element, string property, string value) =>
        (string?)element.Element(MathNamespace + property)?.Element(MathNamespace + value)?.Attribute(MathNamespace + "val");
    private static string[] ReadProperties(XElement element) => DocxSourceXml.StoredChildren(element).Where(IsProperty)
        .Select(child => child.ToString(SaveOptions.DisableFormatting)).ToArray();

    private static DocxMathArgumentProperties? ReadArgumentProperties(XElement element)
    {
        XElement[] properties = element.Elements().Where(IsProperty).ToArray();
        if (properties.Length == 0) { return null; }
        int argumentPropertyCount = properties.Length;
        XElement control = properties[^1];
        if (control.Name == MathNamespace + "ctrlPr")
        {
            // Arguments can carry trailing control character properties. Only
            // controls with no print effect are neutral to the current layout.
            if (element.Name == MathNamespace + "oMath" || control != element.Elements().Last() ||
                !HasKnownControl(control))
                return new(null, true);
            if (control.Element(WordprocessingNamespace + "rPr") is { } runProperties &&
                (runProperties.Elements().Any(property => property.Name.LocalName is not ("noProof" or "webHidden")) ||
                 runProperties.Elements().GroupBy(property => property.Name).Any(group => group.Skip(1).Any())))
                return new(null, true);
            argumentPropertyCount--;
        }
        if (argumentPropertyCount == 0) { return new(null, false); }
        if (argumentPropertyCount != 1 || properties[0].Name != MathNamespace + "argPr")
            return new(null, true);
        XElement argument = properties[0];
        XElement[] children = argument.Elements().ToArray();
        if (argument.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || !HasOnlyWhitespaceText(argument))
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
    internal static bool IsReadableRun(XElement element, bool allowsDeletedText, XElement? firstBreakInRun = null) =>
        element.Name == MathNamespace + "r" && RunContent(element, firstBreakInRun).All(child =>
            ReadLiteralRunText(child, allowsDeletedText) is not null ||
            child.Name == MathNamespace + "rPr" || child.Name == WordprocessingNamespace + "rPr" ||
            DocxSourceXml.IsNonPrintingMarker(child));

    private static IEnumerable<XElement> RunContent(XElement run, XElement? firstBreakInRun) =>
        firstBreakInRun is null ? run.Elements() : run.Elements().TakeWhile(child => child != firstBreakInRun);

    private static string? ReadLiteralRunText(XElement element, bool allowsDeletedText)
    {
        if (element.Name == MathNamespace + "t") return element.Value;
        if (element.HasElements || element.Name.Namespace != WordprocessingNamespace) return null;
        return element.Name.LocalName switch
        {
            "t" => element.Value,
            "delText" when allowsDeletedText => element.Value,
            "noBreakHyphen" or "softHyphen" when string.IsNullOrWhiteSpace(element.Value) &&
                element.Attributes().All(attribute => attribute.IsNamespaceDeclaration) =>
                element.Name.LocalName == "noBreakHyphen" ? "\u2011" : "\u00AD",
            _ => null
        };
    }

    private static DocxMathNode ReadLeaf(XElement element, Func<XElement, string, DocxTextRun> readRun, bool allowsDeletedText,
        XElement? firstBreakInRun)
    {
        if (IsReadableRun(element, allowsDeletedText, firstBreakInRun))
        {
            var (normalText, literalText, supported) = ReadRunProperties(element);
            // Native equations wrap at operator boundaries. Optional break
            // markers contribute neither ink nor advance to the unbroken text.
            string text = string.Concat(RunContent(element, firstBreakInRun).Select(child => ReadLiteralRunText(child, allowsDeletedText)));
            if (text.Any(DocxTextBreakRules.IsOptionalBreakMarker))
                text = new string(text.Where(IsMathLiteralCharacter).ToArray());
            return new DocxMathRun(
                readRun(element, text),
                PropertyValue(element, "rPr", "sty"), PropertyValue(element, "rPr", "scr"))
            {
                PropertiesXml = ReadProperties(element),
                NormalTextValue = normalText,
                LiteralText = literalText,
                ExplicitFontFamily = (string?)element.Element(WordprocessingNamespace + "rPr")?.Element(WordprocessingNamespace + "rFonts")?.Attribute(WordprocessingNamespace + "ascii")
                    ?? (string?)element.Element(WordprocessingNamespace + "rPr")?.Element(WordprocessingNamespace + "rFonts")?.Attribute(WordprocessingNamespace + "hAnsi"),
                HasUnsupportedProperties = !supported
            };
        }
        return new DocxMathUnsupported(element.Name.LocalName, element.Name.NamespaceName, []);
    }

    private static (bool? NormalText, bool LiteralText, bool Supported) ReadRunProperties(XElement run)
    {
        XElement? properties = run.Element(MathNamespace + "rPr");
        var (normalText, normalSupported) = ReadRunFlag(properties, MathNamespace + "nor");
        var (literalText, literalSupported) = ReadRunFlag(properties, MathNamespace + "lit");
        var (alignment, alignmentSupported) = ReadRunFlag(properties, MathNamespace + "aln");
        bool propertiesSupported = properties?.Elements().All(property =>
            property.Name.Namespace == MathNamespace && property.Name.LocalName is "sty" or "scr" or "nor" or "lit" or "aln") != false;
        // Disabled alignment has no layout effect. Active alignment points still need their own placement rule.
        return (normalText, literalText == true, propertiesSupported && normalSupported && literalSupported &&
            alignmentSupported && alignment != true);
    }

    private static (bool? Value, bool Supported) ReadRunFlag(XElement? properties, XName name)
    {
        if (properties?.Element(name) is not { } flag) return (null, true);
        string? source = (string?)flag.Attribute(MathNamespace + "val");
        bool enabled = source is null || OoxBoolean.IsTrue(source);
        bool supported = !properties.Elements(name).Skip(1).Any() && !flag.HasElements &&
            string.IsNullOrWhiteSpace(flag.Value) && (enabled || OoxBoolean.IsOff(source)) &&
            flag.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name == MathNamespace + "val");
        return (enabled, supported);
    }
}
