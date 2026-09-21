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

        foreach (PdfTilingPatternResource pattern in page.Patterns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidatePattern(pattern, pageIndex, cancellationToken);
        }
    }

    private static void ValidatePattern(PdfTilingPatternResource pattern, int pageIndex, CancellationToken cancellationToken)
    {
        string context = $"PDF page {pageIndex + 1} pattern '{pattern.ResourceName}'";
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

    // PLAN G05: tokens retain content offsets instead of substrings. The scan keeps at
    // most four, and only operator classification plus resource-name operands read them:
    // operators classify by length and leading characters, names resolve through span
    // lookups, and strings materialize solely for failure messages. Allocation stays
    // proportional to resource names rather than token count.
    private readonly record struct ContentToken(bool IsName, ContentOperator Operator, int Start, int Length);

    private enum ContentOperator
    {
        Unknown,
        SaveState,
        RestoreState,
        BeginText,
        EndText,
        SetFont,
        DrawImage,
        SetGraphicsState,
        FillShading,
        SetPattern,
    }

    private static ContentOperator ClassifyOperator(string content, int start, int length)
    {
        // Operator characters are ASCII letters plus *, ', " (see IsOperatorChar), so
        // length plus leading characters discriminate the checked set exactly.
        return length switch
        {
            1 => content[start] switch
            {
                'q' => ContentOperator.SaveState,
                'Q' => ContentOperator.RestoreState,
                _ => ContentOperator.Unknown,
            },
            2 when content[start] == 'B' && content[start + 1] == 'T' => ContentOperator.BeginText,
            2 when content[start] == 'E' && content[start + 1] == 'T' => ContentOperator.EndText,
            2 when content[start] == 'T' && content[start + 1] == 'f' => ContentOperator.SetFont,
            2 when content[start] == 'D' && content[start + 1] == 'o' => ContentOperator.DrawImage,
            2 when content[start] == 'g' && content[start + 1] == 's' => ContentOperator.SetGraphicsState,
            2 when content[start] == 's' && content[start + 1] == 'h' => ContentOperator.FillShading,
            3 when content[start] == 's' && content[start + 1] == 'c' && content[start + 2] == 'n' => ContentOperator.SetPattern,
            _ => ContentOperator.Unknown,
        };
    }

    private static void RequireAsciiChar(string content, int position, string context)
    {
        char value = content[position];
        if (value != '\t' && value != '\n' && value != '\r' && (value < ' ' || value > '~'))
        {
            throw new InvalidDataException($"{context} contains a non-ASCII character that ASCII encoding would silently corrupt.");
        }
    }

    private static void RequireAsciiSpan(string content, int start, int end, string context)
    {
        for (int i = start; i < end; i++)
        {
            RequireAsciiChar(content, i, context);
        }
    }

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
                RequireAsciiChar(content, position, context);
                position++;
                continue;
            }

            if (current == '%')
            {
                int commentEnd = position;
                while (commentEnd < content.Length && content[commentEnd] != '\n')
                {
                    commentEnd++;
                }

                RequireAsciiSpan(content, position, commentEnd, context);
                position = commentEnd;
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
                RequireAsciiSpan(content, position + 1, end < 0 ? content.Length : end, context);
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
                int literalEnd = SkipLiteralString(content, position);
                RequireAsciiSpan(content, position, literalEnd, context);
                position = literalEnd;
                continue;
            }

            if (current == '/')
            {
                int start = position + 1;
                int end = start;
                while (end < content.Length && !IsTokenDelimiter(content[end]))
                {
                    RequireAsciiChar(content, end, context);
                    end++;
                }

                PushToken(new ContentToken(IsName: true, ContentOperator.Unknown, start, end - start), recent, ref tokenCount);
                position = end;
                continue;
            }

            if (IsOperatorStart(current))
            {
                int start = position;
                // Operator characters are ASCII by construction (see IsOperatorChar);
                // the leading character was checked above.
                while (position < content.Length && IsOperatorChar(content[position]))
                {
                    position++;
                }

                PushToken(new ContentToken(IsName: false, ClassifyOperator(content, start, position - start), start, position - start), recent, ref tokenCount);
                ApplyOperator(content, recent, context, ref graphicsDepth, ref inText, fonts, images, states, shadings, patterns);
                continue;
            }

            int valueEnd = position;
            while (valueEnd < content.Length && !IsTokenDelimiter(content[valueEnd]) && content[valueEnd] != '/')
            {
                RequireAsciiChar(content, valueEnd, context);
                valueEnd++;
            }

            if (valueEnd != position)
            {
                PushToken(new ContentToken(IsName: false, ContentOperator.Unknown, position, valueEnd - position), recent, ref tokenCount);
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
        string content,
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
        switch (recent[^1].Operator)
        {
            case ContentOperator.SaveState:
                graphicsDepth++;
                return;
            case ContentOperator.RestoreState:
                if (graphicsDepth == 0)
                {
                    throw new InvalidDataException($"{context} restores graphics state without a matching save (stray Q).");
                }

                graphicsDepth--;
                return;
            case ContentOperator.BeginText:
                if (inText)
                {
                    throw new InvalidDataException($"{context} nests text objects (BT inside BT).");
                }

                inText = true;
                return;
            case ContentOperator.EndText:
                if (!inText)
                {
                    throw new InvalidDataException($"{context} ends a text object without opening one (stray ET).");
                }

                inText = false;
                return;
            case ContentOperator.SetFont:
                RequireNameOperand(content, recent, 3, fonts, "font", "Tf", context);
                return;
            case ContentOperator.DrawImage:
                RequireNameOperand(content, recent, 2, images, "image", "Do", context);
                return;
            case ContentOperator.SetGraphicsState:
                RequireNameOperand(content, recent, 2, states, "graphics-state", "gs", context);
                return;
            case ContentOperator.FillShading:
                RequireNameOperand(content, recent, 2, shadings, "shading", "sh", context);
                return;
            case ContentOperator.SetPattern:
                if (recent.Count >= 2 && recent[^2].IsName)
                {
                    RequireNameOperand(content, recent, 2, patterns, "pattern", "scn", context);
                }

                return;
            default:
                return;
        }
    }

    private static void RequireNameOperand(string content, List<ContentToken> recent, int lookback, HashSet<string> names, string category, string op, string context)
    {
        if (recent.Count < lookback || !recent[^lookback].IsName)
        {
            throw new InvalidDataException($"{context} has a malformed {op} operator without a resource name.");
        }

        ContentToken operand = recent[^lookback];
        // Span lookup: resource uses resolve without materializing names. The string
        // form exists only for the failure message below.
        if (!names.GetAlternateLookup<ReadOnlySpan<char>>().Contains(content.AsSpan(operand.Start, operand.Length)))
        {
            string name = content.Substring(operand.Start, operand.Length);
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
