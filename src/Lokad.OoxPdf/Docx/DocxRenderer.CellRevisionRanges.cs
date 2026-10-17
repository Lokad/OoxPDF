namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private sealed record DocxCellRevisionBalloonSource(
        string TableKey, int RowIndex, int GridStart, int GridSpan, DocxRevisionInfo Revision);

    private static IReadOnlyList<DocxMarkupBalloonCandidate> CoalesceCellRevisionRanges(
        IReadOnlyList<DocxMarkupBalloonCandidate> candidates, CancellationToken cancellationToken)
    {
        var result = new List<DocxMarkupBalloonCandidate>(candidates.Count);
        var columns = new Dictionary<(string Table, int Start, int Span), int>();
        foreach (var candidate in candidates)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (candidate.WordCompatibleCellSource is not { } source)
            {
                result.Add(candidate);
                continue;
            }

            var key = (source.TableKey, source.GridStart, source.GridSpan);
            if (columns.TryGetValue(key, out int index))
            {
                var previous = result[index];
                var last = previous.WordCompatibleCellLastSource!;
                // Independent Word controls coalesce consecutive rows in the
                // same column span, including different dates. Same-row and
                // diagonal cells, gaps, different authors and kinds stay apart.
                if (last.RowIndex + 1 == source.RowIndex &&
                    last.Revision.Kind == source.Revision.Kind &&
                    string.Equals(last.Revision.Author, source.Revision.Author, StringComparison.Ordinal))
                {
                    result[index] = previous with
                    {
                        WordCompatibleCellLastSource = source,
                        CandidateCount = previous.CandidateCount + candidate.CandidateCount,
                        RevisionCandidateCount = previous.RevisionCandidateCount + candidate.RevisionCandidateCount
                    };
                    continue;
                }
            }

            columns[key] = result.Count;
            result.Add(candidate);
        }
        return result;
    }
}
