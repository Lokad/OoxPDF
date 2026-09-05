using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private readonly record struct DocxMarkupBalloonRgb(byte Red, byte Green, byte Blue);

    private readonly record struct DocxRevisionMarkupPalette(
        DocxMarkupBalloonRgb FillRgb,
        DocxMarkupBalloonRgb StrokeRgb,
        DocxMarkupBalloonRgb TitleRgb);

    private readonly record struct DocxCommentThreadBalloonMetrics(
        int WithDateCount,
        int ResolvedCount,
        int OpenCount,
        int ReplyCount);

    private sealed record DocxMarkupBalloonArea(
        string Side,
        double X,
        double Width,
        double ConnectorX);

    private sealed record DocxMarkupBalloonCandidate(
        DocxMarkupBalloonKind Kind,
        string Title,
        string Body,
        string? WordCompatibleTitle,
        string? WordCompatibleBody,
        double AnchorY,
        double AnchorConnectorX,
        double AnchorLeftX,
        double AnchorRightX,
        int Sequence,
        DocxMarkupBalloonRgb FillRgb,
        DocxMarkupBalloonRgb StrokeRgb,
        DocxMarkupBalloonRgb TitleRgb,
        DocxMarkupBalloonRgb BodyRgb,
        int CandidateCount,
        int CommentCandidateCount,
        int RevisionCandidateCount,
        int CommentWithDateCount,
        int CommentResolvedCount,
        int CommentOpenCount,
        int CommentReplyCount,
        int BodySummaryPartCount,
        int WordCompatibleBodySummaryPartCount);

    private sealed record DocxMarkupBalloonLaneBand(
        int Index,
        IReadOnlyList<DocxMarkupBalloonCandidate> Candidates,
        double TopLimit,
        double MaxBalloonHeight,
        int CandidateCount);

    private sealed record DocxMarkupBalloonPlacement(
        DocxMarkupBalloonKind Kind,
        string Side,
        string Title,
        string Body,
        string? WordCompatibleTitle,
        string? WordCompatibleBody,
        double X,
        double Y,
        double Width,
        double Height,
        double AnchorY,
        double AnchorConnectorX,
        double BalloonConnectorX,
        bool AnchorConnectorClamped,
        DocxMarkupBalloonRgb FillRgb,
        DocxMarkupBalloonRgb StrokeRgb,
        DocxMarkupBalloonRgb TitleRgb,
        DocxMarkupBalloonRgb BodyRgb,
        bool IsOverflowSummary,
        int CandidateCount,
        int CommentCandidateCount,
        int RevisionCandidateCount,
        int CommentWithDateCount,
        int CommentResolvedCount,
        int CommentOpenCount,
        int CommentReplyCount,
        int BodySummaryPartCount,
        int WordCompatibleBodySummaryPartCount,
        int? OverflowStartIndex,
        int? OverflowEndIndex,
        int LaneBandIndex,
        int LaneBandCandidateCount)
    {
        public DocxMarkupBalloonPlacementSnapshot ToSnapshot(int pageIndex)
        {
            return new DocxMarkupBalloonPlacementSnapshot(
                pageIndex,
                Kind.ToValueString(),
                Side,
                X,
                Y,
                Width,
                Height,
                AnchorY,
                IsOverflowSummary,
                AnchorConnectorX,
                BalloonConnectorX,
                AnchorConnectorClamped,
                CandidateCount,
                CommentCandidateCount,
                RevisionCandidateCount,
                CommentWithDateCount,
                CommentResolvedCount,
                CommentOpenCount,
                CommentReplyCount,
                BodySummaryPartCount,
                WordCompatibleBodySummaryPartCount,
                ResolveCommentThreadSeparatorLineCount(CommentReplyCount),
                OverflowStartIndex,
                OverflowEndIndex,
                LaneBandIndex,
                LaneBandCandidateCount);
        }
    }
}
