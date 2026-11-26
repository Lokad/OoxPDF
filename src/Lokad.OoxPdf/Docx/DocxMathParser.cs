using System.Diagnostics.CodeAnalysis;
using System.Text;
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
        (XElement Run, XElement? FirstBreakInRun)? trailingBreaks = null,
        Func<XElement, bool>? readCellBookmarkBoundary = null)
    {
        var pending = new Stack<(XElement Element, bool Expanded, XName? RevisionOwner, bool IncludeAnnotations, DocxHyperlinkSpan? Hyperlink)>();
        var nodes = new Dictionary<XElement, DocxMathNode>();
        HashSet<XElement>? deferredBookmarkStarts = null;
        IReadOnlySet<(XName Name, string Id)>? moveOwners = null;
        XName deletion = WordprocessingNamespace + "del";
        pending.Push((equation, false, null, true, null));
        while (pending.TryPop(out var item))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = item.Element;
            string name = element.Name.LocalName;
            if (DocxSourceXml.IsCommentBoundary(element))
            {
                if (item.IncludeAnnotations) readCommentBoundary?.Invoke(element);
                continue;
            }
            if (DocxSourceXml.IsBookmarkMarker(element))
            {
                if (item.IncludeAnnotations && readCellBookmarkBoundary?.Invoke(element) == true)
                    (deferredBookmarkStarts ??= []).Add(element);
                continue;
            }
            bool revisionSource = item.RevisionOwner is null && includeRevision is not null &&
                DocxSourceXml.IsAdmittedMathRevisionOwner(element, equation, authoredStory, ref moveOwners, cancellationToken);
            XElement[]? revisionChildren = revisionSource ? ReadRevisionChildren(element, cancellationToken).ToArray() : null;
            bool deleting = element.Name == deletion;
            bool revisionRow = revisionChildren is not null && IsReadableRevision(revisionChildren, deleting);
            bool included = !revisionRow || includeRevision!(element);
            bool retainDeletedFrame = revisionRow && preserveDeletedStructure && deleting && !included &&
                revisionChildren!.Any(child => IsMathControl(child.Name));
            if (revisionRow && !item.Expanded && !included && !retainDeletedFrame)
            {
                nodes.Add(element, new DocxMathRow([]));
                continue;
            }
            bool sourceContainer = IsSourceContainer(element);
            bool phantom = element.Name == MathNamespace + "phant";
            bool unsupportedPhantom = phantom && !item.Expanded &&
                (!HasOnlyRoles(element, "e") || !IsVisibleTransparentPhantom(element, cancellationToken));
            if (!revisionRow && !sourceContainer && (!IsContainer(element.Name) || unsupportedPhantom))
            {
                bool allowsDeletedText = item.RevisionOwner is null ? equationHasDeletionOwner : item.RevisionOwner == deletion;
                XElement? firstBreakInRun = element == trailingBreaks?.Run ? trailingBreaks?.FirstBreakInRun : null;
                var readRevision = item.RevisionOwner is null ? includeRevision : null;
                nodes.Add(element, ReadLeaf(element, readRun, allowsDeletedText, firstBreakInRun, readRevision,
                    item.IncludeAnnotations ? item.Hyperlink : null, cancellationToken));
                continue;
            }

            bool mathRow = IsRow(name) && (name == "oMath" || !IsSourceContainer(element.Parent)) &&
                (name != "deg" || element.Parent?.Name == MathNamespace + "rad") &&
                (name != "fName" || element.Parent?.Name == MathNamespace + "func");
            if (!item.Expanded)
            {
                DocxHyperlinkSpan? hyperlink = item.IncludeAnnotations && included ? item.Hyperlink : null;
                if (item.IncludeAnnotations && included && element.Name == WordprocessingNamespace + "hyperlink")
                    hyperlink = readHyperlink?.Invoke(element);
                pending.Push((element, true, item.RevisionOwner, item.IncludeAnnotations && included, hyperlink));
                XName? revisionOwner = revisionRow ? element.Name : item.RevisionOwner;
                // Range boundaries can cross nested rows. Visit them in source
                // order before using the rows to assemble math geometry.
                bool visitCellBookmarks = readCellBookmarkBoundary is not null && (revisionRow || sourceContainer || mathRow || phantom);
                foreach (XElement child in WithoutTrailingBreaks(element, TraversalChildren(element, visitCellBookmarks)).Reverse())
                {
                    pending.Push((child, false, revisionOwner,
                        item.IncludeAnnotations && included, hyperlink));
                }
                continue;
            }

            DocxMathNode node;
            if (revisionRow || sourceContainer)
                node = new DocxMathRow(ContentChildren(element).Select(child => nodes[child]).ToArray());
            else if (mathRow)
            {
                node = new DocxMathRow(WithoutTrailingBreaks(element, ContentChildren(element)).Select(child => nodes[child]).ToArray())
                { ArgumentProperties = ReadArgumentProperties(element) };
            }
            else if (phantom && Role(element, "e") is { } visibleBase)
            {
                node = new DocxMathRow([visibleBase]);
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
                    if (DocxSourceXml.IsBookmarkMarker(child))
                    {
                        if (child.Name.LocalName == "bookmarkStart" && deferredBookmarkStarts?.Contains(child) != true)
                        {
                            var bookmark = new DocxMathBookmark((string?)child.Attribute(WordprocessingNamespace + "id"),
                                (string?)child.Attribute(WordprocessingNamespace + "name"));
                            (bookmarks ??= []).Add(new(bookmark, childIndex));
                            readBookmark?.Invoke(bookmark);
                        }
                    }
                    else if (!IsProperty(child) && !IsNonPrintingSourceMarker(child) && !DocxSourceXml.IsCommentBoundary(child)) childIndex++;
                }
                if (bookmarks is not null) node = sourceRow with { Bookmarks = bookmarks };
            }
            node = node with { Hyperlink = item.IncludeAnnotations && included ? item.Hyperlink : null };
            nodes.Add(element, node);
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
    internal static bool IsMathRow(XName? name) => name is not null && name.Namespace == MathNamespace && IsRow(name.LocalName);
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
                _ => IsPrintNeutralControlProperty(property)
            });
    }

    private static bool IsPrintNeutralControlProperty(XElement property)
    {
        string? value = (string?)property.Attribute(WordprocessingNamespace + "val");
        bool neutral = property.Name.LocalName switch
        {
            "noProof" or "webHidden" => HasBooleanValue(value),
            "outline" or "shadow" or "emboss" or "imprint" => OoxBoolean.IsOff(value),
            "effect" or "em" or "u" or "highlight" => value == "none",
            "bdr" => value is "none" or "nil",
            _ => false
        };
        return neutral && property.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == WordprocessingNamespace + "val");
    }
    // A single control has the same grammar inside an admitted revision wrapper.
    // Role rows and orphan matrix rows still require their actual source parent.
    internal static bool IsMathControl(XName name) => name.Namespace == MathNamespace &&
        name.LocalName is "f" or "sSub" or "sSup" or "sSubSup" or "rad" or "d" or "nary" or "func" or "m" or "phant";

    private static bool IsReadableRevision(XElement[] children, bool allowsDeletedText)
    {
        // A revision needs an actual operand or structural control.
        // Readable zero-content siblings do not invalidate that owner.
        bool hasContent = false;
        foreach (var child in children)
        {
            if (IsRevisionContent(child, allowsDeletedText))
                hasContent = true;
            else if (!IsReadableRun(child, allowsDeletedText) || !ReadRunProperties(child.Element(MathNamespace + "rPr")).Supported)
                return false;
        }
        return hasContent;
    }

    private static bool IsRevisionContent(XElement element, bool allowsDeletedText) => element.Name.Namespace == MathNamespace &&
        (IsMathControl(element.Name) || element.Name.LocalName == "r" &&
            HasMathRunText(element.Elements(), allowsDeletedText));

    private static bool HasMathRunText(IEnumerable<XElement> children, bool allowsDeletedText) =>
        children.Any(child => ReadLiteralRunText(child, allowsDeletedText, out _)?.Any(IsMathLiteralCharacter) == true);

    private static bool IsMathLiteralCharacter(char value) => !OoxTextBreakRules.IsNonPrintingBreakControl(value);
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
    // Directional wrappers preserve stored math in source order. The source
    // probe diagnoses the approximation separately from the projected content.
    internal static bool IsMathStoredContentContainer(XElement element) =>
        DocxSourceXml.IsInlineContentContainer(element) &&
        (element.Name != WordprocessingNamespace + "sdtContent" || element.Parent?.Name == WordprocessingNamespace + "sdt");

    private static bool IsSourceContainer(XElement? element) =>
        element is not null && (IsMathStoredContentContainer(element) || element.Name == WordprocessingNamespace + "hyperlink");
    private static bool IsContainer(XName name) => IsMathRow(name) || IsMathControl(name) || name == MathNamespace + "mr";
    private static bool IsProperty(XElement element) => element.Name.Namespace == MathNamespace && element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal);
    internal static bool IsNonPrintingSourceMarker(XElement element) =>
        DocxSourceXml.IsNonPrintingMarker(element) ||
        DocxSourceXml.TryResolveRevisionRangeMarker(element, out _, out _) &&
        DocxSourceXml.IsNonPrintingRevisionRangeBoundary(element);

    private static IEnumerable<XElement> TraversalChildren(XElement element, bool visitCellBookmarks) =>
        DocxSourceXml.StoredChildren(element).Where(child => !IsProperty(child) && !IsNonPrintingSourceMarker(child) &&
            (visitCellBookmarks || !DocxSourceXml.IsBookmarkMarker(child)));
    private static IEnumerable<XElement> NonMarkerChildren(XElement element) =>
        DocxSourceXml.StoredChildren(element).Where(child => !IsNonPrintingSourceMarker(child) && !DocxSourceXml.IsBookmarkMarker(child) && !DocxSourceXml.IsCommentBoundary(child));
    private static IEnumerable<XElement> ContentChildren(XElement element) =>
        NonMarkerChildren(element).Where(child => !IsProperty(child));
    private static bool HasOnlyRoles(XElement element, params string[] roles)
    {
        XElement[] children = ContentChildren(element).ToArray();
        return children.Length == roles.Length && children.All(child => child.Name.Namespace == MathNamespace && roles.Contains(child.Name.LocalName));
    }
    private static string? PropertyValue(XElement? properties, string value) =>
        (string?)properties?.Element(MathNamespace + value)?.Attribute(MathNamespace + "val");
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
                !HasPrintNeutralControl(control))
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
        element.Name == MathNamespace + "r" && HasReadableRunChildren(RunContent(element, firstBreakInRun), allowsDeletedText);

    internal static bool HasOnlyDeletedLiteralRuns(XElement equation, CancellationToken token) =>
        ReadRevisionChildren(equation, token).All(child => IsReadableRun(child, allowsDeletedText: true));

    private static bool HasReadableRunChildren(IEnumerable<XElement> children, bool allowsDeletedText) =>
        children.All(child =>
            ReadLiteralRunText(child, allowsDeletedText, out _) is not null ||
            child.Name == MathNamespace + "rPr" || child.Name == WordprocessingNamespace + "rPr" ||
            DocxSourceXml.IsNonPrintingMarker(child));

    private static IEnumerable<XElement> RunContent(XElement run, XElement? firstBreakInRun) =>
        firstBreakInRun is null ? run.Elements() : run.Elements().TakeWhile(child => child != firstBreakInRun);

    private static string? ReadLiteralRunText(XElement element, bool allowsDeletedText, out string? symbolFont)
    {
        symbolFont = null;
        if (element.Name == MathNamespace + "t") return element.Value;
        if (element.HasElements || element.Name.Namespace != WordprocessingNamespace) return null;
        if (TryReadSymbol(element, out symbolFont, out char symbol)) return symbol.ToString();
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

    internal static bool TryReadSymbol(XElement element, [NotNullWhen(true)] out string? font, out char symbol)
    {
        font = null;
        symbol = default;
        return element.Name == WordprocessingNamespace + "sym" && !element.HasElements &&
            string.IsNullOrWhiteSpace(element.Value) &&
            element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
                attribute.Name == WordprocessingNamespace + "font" || attribute.Name == WordprocessingNamespace + "char") &&
            DocxReader.TryReadSymbol(element, out font, out symbol);
    }

    private static XElement? ReadRunRevision(XElement run, CancellationToken token)
    {
        // The revision owns all literal text. Each property kind can be kept
        // on the enclosing run or packed inside the revision with the text.
        // Nonprinting markers can surround the revision without owning text.
        XElement? mathProperties = null;
        XElement? wordProperties = null;
        XElement? revision = null;
        foreach (XElement child in run.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (DocxSourceXml.IsNonPrintingMarker(child)) continue;
            if (revision is not null) return null;

            if (child.Name == MathNamespace + "rPr" && mathProperties is null)
                mathProperties = child;
            else if (child.Name == WordprocessingNamespace + "rPr" && wordProperties is null)
                wordProperties = child;
            else if (DocxSourceXml.IsPackedMathRunRevision(child))
                revision = child;
            else
                return null;
        }

        return revision is not null &&
            (mathProperties is null || revision.Element(mathProperties.Name) is null) &&
            (wordProperties is null || revision.Element(wordProperties.Name) is null) &&
            HasOnlyWhitespaceText(run) && HasOnlyWhitespaceText(revision) ? revision : null;
    }

    internal static XElement? ReadWordRunProperties(XElement source) =>
        ReadRunProperty(source, WordprocessingNamespace + "rPr");

    private static XElement? ReadRunProperty(XElement source, XName name) =>
        source.Element(name) ?? (DocxSourceXml.IsPackedMathRunRevision(source) ? source.Parent?.Element(name) : null);

    private static DocxMathNode ReadLeaf(XElement element, Func<XElement, string, DocxTextRun> readRun, bool allowsDeletedText,
        XElement? firstBreakInRun, Func<XElement, bool>? includeRevision, DocxHyperlinkSpan? hyperlink, CancellationToken token)
    {
        if (element.Name != MathNamespace + "r") return new DocxMathUnsupported(element.Name.LocalName, element.Name.NamespaceName, []);
        XElement? revision = includeRevision is not null && firstBreakInRun is null ? ReadRunRevision(element, token) : null;
        XElement owner = revision ?? element;
        if (revision is not null) allowsDeletedText = revision.Name == WordprocessingNamespace + "del";
        var children = RunContent(owner, firstBreakInRun);
        if (!HasReadableRunChildren(children, allowsDeletedText) || revision is not null && !HasMathRunText(children, allowsDeletedText))
            return new DocxMathUnsupported(element.Name.LocalName, element.Name.NamespaceName, []);
        if (revision is not null && !includeRevision!(revision)) return new DocxMathRow([]);

        List<(int Offset, string Font)>? symbols = null;
        int offset = 0;
        string text = string.Concat(children.Select(child =>
        {
            token.ThrowIfCancellationRequested();
            string? value = ReadLiteralRunText(child, allowsDeletedText, out string? font);
            if (font is not null) (symbols ??= []).Add((offset, font));
            offset += value?.Length ?? 0;
            return value;
        }));
        XElement? mathProperties = ReadRunProperty(owner, MathNamespace + "rPr");
        var (normalText, literalText, supported) = ReadRunProperties(mathProperties, text);
        var fonts = ReadWordRunProperties(owner)?.Element(WordprocessingNamespace + "rFonts");
        string? family = (string?)fonts?.Attribute(WordprocessingNamespace + "ascii")
            ?? (string?)fonts?.Attribute(WordprocessingNamespace + "hAnsi");
        string? style = PropertyValue(mathProperties, "sty");
        string? script = PropertyValue(mathProperties, "scr");
        DocxTextRun? resolvedRun = null;
        if (symbols is null) return CreateRun(text, 0);

        // A symbol selects one font-relative glyph inside the authored run.
        // Keep neighboring mathematical text in its original style and order.
        var fragments = new List<DocxMathNode>();
        int start = 0;
        foreach (var (position, font) in symbols)
        {
            token.ThrowIfCancellationRequested();
            if (position > start) fragments.Add(CreateRun(text[start..position], start));
            fragments.Add(CreateRun(text.Substring(position, 1), position, font));
            start = position + 1;
        }
        if (start < text.Length) fragments.Add(CreateRun(text[start..], start));
        return fragments.Count == 1 ? fragments[0] : new DocxMathRow(fragments);

        DocxMathRun CreateRun(string value, int sourceOffset, string? symbolFont = null)
        {
            // Optional breaks have no ink. Word joiners retain their boundary
            // even if a symbol splits the source run into several fragments.
            List<int>? joinerOffsets = null;
            if (value.Any(character => !IsMathLiteralCharacter(character)))
            {
                var literal = new StringBuilder(value.Length);
                foreach (char character in value)
                {
                    token.ThrowIfCancellationRequested();
                    if (IsMathLiteralCharacter(character))
                        literal.Append(character);
                    else if (OoxTextBreakRules.IsWordJoiner(character) &&
                        (joinerOffsets is null || joinerOffsets[^1] != literal.Length))
                        (joinerOffsets ??= []).Add(literal.Length);
                }
                value = literal.ToString();
            }

            var run = resolvedRun ??= readRun(owner, value);
            if (symbolFont is not null)
            {
                run = run with { FontFamily = symbolFont, Fonts = new(symbolFont, symbolFont, symbolFont, symbolFont, null, null, null, null),
                    AllCaps = false, SmallCaps = false };
            }
            if (sourceOffset != 0 || run.Text != value)
                run = run with { Text = value, SourceTextOffsetInRun = run.SourceTextOffsetInRun + sourceOffset };
            return new DocxMathRun(run, style, script)
            {
                Hyperlink = hyperlink,
                WordJoinerOffsets = joinerOffsets is null ? [] : joinerOffsets,
                NormalTextValue = symbolFont is not null ? true : normalText,
                LiteralText = literalText,
                ExplicitFontFamily = symbolFont ?? family,
                HasUnsupportedProperties = !supported
            };
        }
    }

    private static (bool? NormalText, bool LiteralText, bool Supported) ReadRunProperties(XElement? properties, string text = "")
    {
        var (normalText, normalSupported) = ReadRunFlag(properties, MathNamespace + "nor");
        var (literalText, literalSupported) = ReadRunFlag(properties, MathNamespace + "lit");
        var (alignment, alignmentSupported) = ReadRunFlag(properties, MathNamespace + "aln");
        XElement? manualBreak = properties?.Element(MathNamespace + "brk");
        bool ignoresBreak = manualBreak is not null && IsIgnoredManualRunBreak(manualBreak, text);
        bool propertiesSupported = properties?.Elements().All(property =>
            property.Name.Namespace == MathNamespace &&
            (property.Name.LocalName is "sty" or "scr" or "nor" or "lit" or "aln" || ignoresBreak && property == manualBreak)) != false;
        // Disabled alignment has no layout effect. Active alignment points still need their own placement rule.
        return (normalText, literalText == true, propertiesSupported && normalSupported && literalSupported &&
            alignmentSupported && alignment != true);
    }

    private static bool IsIgnoredManualRunBreak(XElement instruction, string text)
    {
        // Word ignores m:brk when the run does not begin with an operator.
        // A leading letter or digit is a literal operand. Keep the fallback
        // for operator-led runs and other starts whose interpretation is unknown.
        XName alignAt = MathNamespace + "alnAt";
        int indexOfFirst = 0;
        while (indexOfFirst < text.Length && !IsMathLiteralCharacter(text[indexOfFirst])) indexOfFirst++;
        return indexOfFirst < text.Length && Rune.TryGetRuneAt(text, indexOfFirst, out Rune first) && Rune.IsLetterOrDigit(first) &&
            !instruction.HasElements && string.IsNullOrWhiteSpace(instruction.Value) &&
            instruction.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name == alignAt) &&
            ((string?)instruction.Attribute(alignAt) is not { } value ||
                int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture,
                    out int index) && index is >= 1 and <= 255);
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
