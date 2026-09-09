namespace Lokad.OoxPdf.Pdf;

// Independent structural gate over emitted pages: the graphics builder trusts its
// callers for resource references and graphics/text state, so the writer verifies
// each page before writing any object. The scan covers exactly the operator set
// the builder emits (no inline images, no content dictionaries); anything outside
// that set still tokenizes, but unknown operators carry no checks.
// Tiling-pattern streams get the same treatment against their own image sets,
// since pattern content is caller-supplied free text; soft-mask streams are
// generated next to their resources and stay consistent by construction.
internal static class PdfContentValidator
{
    public static void ValidatePage(PdfPage page, int pageIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        RequireUniqueNames(page.Fonts.Select(font => font.ResourceName), "font", pageIndex);
        RequireUniqueNames(page.Images.Select(image => image.ResourceName), "image", pageIndex);
        RequireUniqueNames(page.ExtGStates.Select(state => state.ResourceName), "graphics-state", pageIndex);
        RequireUniqueNames(page.Shadings.Select(shading => shading.ResourceName), "shading", pageIndex);
        RequireUniqueNames(page.Patterns.Select(pattern => pattern.ResourceName), "pattern", pageIndex);

        ValidateContent(
            page.Content,
            $"PDF page {pageIndex + 1}",
            new HashSet<string>(page.Fonts.Select(font => PdfEmbeddedFont.SanitizeName(font.ResourceName)), StringComparer.Ordinal),
            new HashSet<string>(page.Images.Select(image => PdfEmbeddedFont.SanitizeName(image.ResourceName)), StringComparer.Ordinal),
            new HashSet<string>(page.ExtGStates.Select(state => PdfEmbeddedFont.SanitizeName(state.ResourceName)), StringComparer.Ordinal),
            new HashSet<string>(page.Shadings.Select(shading => PdfEmbeddedFont.SanitizeName(shading.ResourceName)), StringComparer.Ordinal),
            new HashSet<string>(page.Patterns.Select(pattern => PdfEmbeddedFont.SanitizeName(pattern.ResourceName)), StringComparer.Ordinal),
            cancellationToken);

        RequireAscii(page.Content, $"PDF page {pageIndex + 1} content");
        foreach (PdfTilingPatternResource pattern in page.Patterns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidatePattern(pattern, pageIndex, cancellationToken);
        }
    }

    private static void ValidatePattern(PdfTilingPatternResource pattern, int pageIndex, CancellationToken cancellationToken)
    {
        string context = $"PDF page {pageIndex + 1} pattern '{pattern.ResourceName}'";
        RequireAscii(pattern.Pattern.Content, context);
        ValidateContent(
            pattern.Pattern.Content,
            context,
            new HashSet<string>(StringComparer.Ordinal),
            new HashSet<string>(pattern.Pattern.Images.Select(image => PdfEmbeddedFont.SanitizeName(image.ResourceName)), StringComparer.Ordinal),
            new HashSet<string>(StringComparer.Ordinal),
            new HashSet<string>(StringComparer.Ordinal),
            new HashSet<string>(StringComparer.Ordinal),
            cancellationToken);
    }

    private static void RequireAscii(string content, string context)
    {
        foreach (char c in content)
        {
            if (c != '\t' && c != '\n' && c != '\r' && (c < ' ' || c > '~'))
            {
                throw new InvalidDataException($"{context} contains a non-ASCII character that ASCII encoding would silently corrupt.");
            }
        }
    }

    private static void RequireUniqueNames(IEnumerable<string> resourceNames, string category, int pageIndex)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (string resourceName in resourceNames)
        {
            string sanitized = PdfEmbeddedFont.SanitizeName(resourceName);
            if (!seen.Add(sanitized))
            {
                throw new InvalidDataException($"PDF page {pageIndex + 1} has a duplicate {category} resource name '{sanitized}'.");
            }
        }
    }

    private readonly record struct ContentToken(bool IsName, bool IsOperator, string Text);

    private static void ValidateContent(
        string content,
        string context,
        HashSet<string> fonts,
        HashSet<string> images,
        HashSet<string> states,
        HashSet<string> shadings,
        HashSet<string> patterns,
        CancellationToken cancellationToken)
    {
        int graphicsDepth = 0;
        bool inText = false;
        // Rolling window of recent tokens for operand lookback (Tf reads two back).
        var recent = new List<ContentToken>(4);
        int tokenCount = 0;
        int position = 0;
        while (position < content.Length)
        {
            if ((tokenCount & 1023) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            char current = content[position];
            if (char.IsWhiteSpace(current))
            {
                position++;
                continue;
            }

            if (current == '%')
            {
                while (position < content.Length && content[position] != '\n')
                {
                    position++;
                }

                continue;
            }

            if (current == '<')
            {
                if (position + 1 < content.Length && content[position + 1] == '<')
                {
                    position += 2;
                    continue;
                }

                int end = content.IndexOf('>', position + 1);
                position = end < 0 ? content.Length : end + 1;
                continue;
            }

            if (current == '>')
            {
                position += position + 1 < content.Length && content[position + 1] == '>' ? 2 : 1;
                continue;
            }

            if (current == '(')
            {
                position = SkipLiteralString(content, position);
                continue;
            }

            if (current == '/')
            {
                int start = position + 1;
                int end = start;
                while (end < content.Length && !IsTokenDelimiter(content[end]))
                {
                    end++;
                }

                PushToken(new ContentToken(IsName: true, IsOperator: false, content.Substring(start, end - start)), recent, ref tokenCount);
                position = end;
                continue;
            }

            if (IsOperatorStart(current))
            {
                int start = position;
                while (position < content.Length && IsOperatorChar(content[position]))
                {
                    position++;
                }

                PushToken(new ContentToken(IsName: false, IsOperator: true, content.Substring(start, position - start)), recent, ref tokenCount);
                ApplyOperator(recent, context, ref graphicsDepth, ref inText, fonts, images, states, shadings, patterns);
                continue;
            }

            int valueEnd = position;
            while (valueEnd < content.Length && !IsTokenDelimiter(content[valueEnd]) && content[valueEnd] != '/')
            {
                valueEnd++;
            }

            if (valueEnd != position)
            {
                PushToken(new ContentToken(IsName: false, IsOperator: false, content.Substring(position, valueEnd - position)), recent, ref tokenCount);
            }
            position = valueEnd == position ? position + 1 : valueEnd;
        }

        if (graphicsDepth != 0)
        {
            throw new InvalidDataException($"{context} has unbalanced graphics state (q/Q depth {graphicsDepth}).");
        }

        if (inText)
        {
            throw new InvalidDataException($"{context} leaves a text object open (BT without ET).");
        }
    }

    private static void PushToken(ContentToken token, List<ContentToken> recent, ref int tokenCount)
    {
        recent.Add(token);
        if (recent.Count > 4)
        {
            recent.RemoveAt(0);
        }

        tokenCount++;
    }

    private static void ApplyOperator(
        List<ContentToken> recent,
        string context,
        ref int graphicsDepth,
        ref bool inText,
        HashSet<string> fonts,
        HashSet<string> images,
        HashSet<string> states,
        HashSet<string> shadings,
        HashSet<string> patterns)
    {
        string op = recent[^1].Text;
        switch (op)
        {
            case "q":
                graphicsDepth++;
                return;
            case "Q":
                if (graphicsDepth == 0)
                {
                    throw new InvalidDataException($"{context} restores graphics state without a matching save (stray Q).");
                }

                graphicsDepth--;
                return;
            case "BT":
                if (inText)
                {
                    throw new InvalidDataException($"{context} nests text objects (BT inside BT).");
                }

                inText = true;
                return;
            case "ET":
                if (!inText)
                {
                    throw new InvalidDataException($"{context} ends a text object without opening one (stray ET).");
                }

                inText = false;
                return;
            case "Tf":
                RequireNameOperand(recent, 3, fonts, "font", "Tf", context);
                return;
            case "Do":
                RequireNameOperand(recent, 2, images, "image", "Do", context);
                return;
            case "gs":
                RequireNameOperand(recent, 2, states, "graphics-state", "gs", context);
                return;
            case "sh":
                RequireNameOperand(recent, 2, shadings, "shading", "sh", context);
                return;
            case "scn":
                if (recent.Count >= 2 && recent[^2].IsName)
                {
                    RequireNameOperand(recent, 2, patterns, "pattern", "scn", context);
                }

                return;
            default:
                return;
        }
    }

    private static void RequireNameOperand(List<ContentToken> recent, int lookback, HashSet<string> names, string category, string op, string context)
    {
        if (recent.Count < lookback || !recent[^lookback].IsName)
        {
            throw new InvalidDataException($"{context} has a malformed {op} operator without a resource name.");
        }

        string name = recent[^lookback].Text;
        if (!names.Contains(name))
        {
            throw new InvalidDataException($"{context} references an unlisted {category} resource '{name}' in {op}.");
        }
    }

    private static int SkipLiteralString(string content, int position)
    {
        int depth = 0;
        while (position < content.Length)
        {
            char current = content[position];
            if (current == '\\' && position + 1 < content.Length)
            {
                position += 2;
                continue;
            }

            if (current == '(')
            {
                depth++;
            }
            else if (current == ')')
            {
                depth--;
                if (depth == 0)
                {
                    return position + 1;
                }
            }

            position++;
        }

        return position;
    }

    private static bool IsTokenDelimiter(char current)
    {
        return char.IsWhiteSpace(current) || current is '(' or ')' or '<' or '>' or '[' or ']' or '{' or '}' or '/' or '%';
    }

    private static bool IsOperatorStart(char current)
    {
        return (current >= 'A' && current <= 'Z') || (current >= 'a' && current <= 'z') || current == '*' || current == '\'' || current == '"';
    }

    private static bool IsOperatorChar(char current)
    {
        return IsOperatorStart(current);
    }
}
