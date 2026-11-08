using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxWordCompatibleMixedBalloonSplitsContinuationWordsInTailFace()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (double rightMargin in new[] { 72d, 207d })
        {
            const string prefix = "Review ";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            string tail = "ok. " + word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "OtherCommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(900, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            var tailRows = shows.Skip(1).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 3, "Overwide continuation words must split into source-face rows.");
            TestAssert.True(tailRows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('a')) == 700),
                "Every word fragment must retain the tail face.");
            for (int i = 0; i < tailRows.Length; i++)
            {
                TestAssert.True(tailRows[i].X + tailRows[i].Font.MeasureTextPoints(tailRows[i].Text.TrimEnd(), tailRows[i].Size)
                    <= balloon.X + balloon.Width + 0.02d, "Word fragments must stay inside the balloon.");
                if (i > 0)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - 1.45d * tailRows[i].Size) < 0.02d,
                        "Continuation pitch must use the actual tail face.");
                }
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - tailRows[^1].Size) < 0.001d, "The final mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + tailRows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Mixed balloon height must include every word-fragment row.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonSplitsJoinedFirstWordInSourceFaces()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (double rightMargin in new[] { 72d, 207d })
        {
            const string prefix = "Review";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            string tail = word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "OtherCommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(900, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            var tailRows = shows.Skip(1).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 3, "The joined first word must have multiple fragments.");
            TestAssert.True(Math.Abs(tailRows[0].Y - shows[0].Y) < 0.001d, "The first fragment stays beside its joined prefix.");
            double tailX = shows[0].X + shows[0].Font.MeasureTextPoints(prefix, shows[0].Size);
            TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "The joined tail follows its prefix face advance.");
            for (int i = 0; i < tailRows.Length; i++)
            {
                TestAssert.Equal(700, (int)tailRows[i].Font.Font.GetAdvanceWidth(tailRows[i].Font.Font.MapCodePoint('a')));
                TestAssert.True(tailRows[i].X + tailRows[i].Font.MeasureTextPoints(tailRows[i].Text.TrimEnd(), tailRows[i].Size)
                    <= balloon.X + balloon.Width + 0.02d, "Joined-word fragments must stay inside the balloon.");
                if (i > 0)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - 1.45d * tailRows[i].Size) < 0.02d,
                        "The joined word's continuation pitch uses its tail face.");
                }
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - tailRows[^1].Size) < 0.001d, "The final mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + tailRows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Height must include the joined word's source-face fragments.");

            DocxParagraph composed = comment with { Runs = [comment.Runs[0] with { Text = "Review table," }, comment.Runs[1]] };
            DocxDocument fallback = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(composed)] }] };
            var fallbackBalloon = renderer.InspectMarkupBalloons(fallback).Single();
            var fallbackShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(fallback, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= fallbackBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                    !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.True(fallbackShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                "A punctuated multi-word prefix retains the complete fallback path.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonSeparatesFirstFontTransitionFromTailPitch()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            const string prefix = "Review ";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            string tail = word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "OtherCommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodyDescender: -350),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(900, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            var tailRows = shows.Skip(1).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 3, "The separated overwide word must start below the prefix and split.");
            double size = tailRows[0].Size;
            TestAssert.True(Math.Abs(shows[0].Y - tailRows[0].Y - 1.45d * size) < 0.02d,
                "The first step uses prefix descent, tail ascent and tail line gap.");
            for (int i = 0; i < tailRows.Length; i++)
            {
                TestAssert.Equal(700, (int)tailRows[i].Font.Font.GetAdvanceWidth(tailRows[i].Font.Font.MapCodePoint('a')));
                TestAssert.True(tailRows[i].X + tailRows[i].Font.MeasureTextPoints(tailRows[i].Text.TrimEnd(), size)
                    <= balloon.X + balloon.Width + 0.02d, "Separated-word fragments must stay inside the balloon.");
                if (i > 0)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - 1.60d * size) < 0.02d,
                        "Later steps retain the tail face's own descent and line height.");
                }
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The final mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Box height must include the distinct first transition and later tail steps.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonComposesJoinedPrefixAcrossFontRows()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (bool boundary in new[] { false, true })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            const string prefix = "Review table";
            int scalarCount = boundary ? 14 : 72;
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", scalarCount / 2)) : new string('a', scalarCount);
            string tail = word + (boundary ? " " + new string('b', 8) : string.Empty) + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodyDescender: -350, bodyAscender: 500),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal("Review ", shows[0].Text);
            TestAssert.Equal("table", shows[1].Text);
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            TestAssert.True(shows.Take(2).All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 700),
                "Both prefix rows must retain the prepared prefix face.");
            var tailRows = shows.Skip(2).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= (boundary ? 2 : 3), "The composed tail must continue below its mixed row.");
            double size = tailRows[0].Size;
            TestAssert.True(Math.Abs(shows[0].Y - shows[1].Y - 1.30d * size) < 0.02d,
                "The first step uses prefix descent plus the larger complete ascent/line-gap metric.");
            TestAssert.True(Math.Abs(shows[1].Y - tailRows[0].Y) < 0.001d, "The final prefix word joins the first tail row.");
            double tailX = shows[1].X + shows[1].Font.MeasureTextPoints("table", size);
            TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "The tail starts after the residual prefix advance.");
            if (boundary)
            {
                TestAssert.Equal(word + " " + new string('b', 8) + " ", tailRows[0].Text);
            }
            for (int i = 0; i < tailRows.Length; i++)
            {
                TestAssert.Equal(900, (int)tailRows[i].Font.Font.GetAdvanceWidth(tailRows[i].Font.Font.MapCodePoint('a')));
                TestAssert.True(tailRows[i].X + tailRows[i].Font.MeasureTextPoints(tailRows[i].Text.TrimEnd(), size)
                    <= balloon.X + balloon.Width + 0.02d, "Visible tail words must stay inside the balloon.");
                if (i > 0)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - (i == 1 ? 1.15d : 1d) * size) < 0.02d,
                        "The mixed descent applies once; later rows retain the tail's own metrics.");
                }
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Balloon height includes both font transitions and every later tail row.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsLeadingTailSeparatorInSourceFace()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        foreach (string separator in new[] { " ", "  " })
        {
            const string prefix = "Review table";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            string tail = separator + "ok. " + word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodySpaceAdvance: 123,
                bodyDescender: -350, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            var tailRows = shows.Skip(1).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 4, "The leading separator must coexist with later within-word splits.");
            TestAssert.Equal(separator + "ok. ", tailRows[0].Text);
            double size = tailRows[0].Size;
            TestAssert.True(Math.Abs(shows[0].Y - tailRows[0].Y) < 0.001d, "The fitting tail shares the prefix row.");
            double tailX = shows[0].X + shows[0].Font.MeasureTextPoints(prefix, size);
            TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "The leading separator starts after the prefix advance.");
            TestAssert.Equal(900, (int)tailRows[0].Font.Font.GetAdvanceWidth(tailRows[0].Font.Font.MapCodePoint(' ')));
            for (int i = 0; i < tailRows.Length; i++)
            {
                TestAssert.Equal(900, (int)tailRows[i].Font.Font.GetAdvanceWidth(tailRows[i].Font.Font.MapCodePoint('a')));
                TestAssert.True(tailRows[i].X + tailRows[i].Font.MeasureTextPoints(tailRows[i].Text.TrimEnd(), size)
                    <= balloon.X + balloon.Width + 0.02d, "Visible tail words must stay inside the balloon.");
                if (i > 0)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - (i == 1 ? 1.15d : 1d) * size) < 0.02d,
                        "The first step uses mixed descent; later steps use the tail face alone.");
                }
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The final mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Height includes the mixed first step and later tail rows.");

            foreach (string guardedTail in new[] { " ".PadRight(64) + "ok. " + word + "  ending." })
            {
                DocxParagraph guardedComment = comment with { Runs = [comment.Runs[0], comment.Runs[1] with { Text = guardedTail }] };
                DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                        !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "A leading space span with repeated internal spaces retains fallback.");
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonSeparatesLeadingFirstWordInItsSourceFace()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (bool zeroAdvance in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (int length in new[] { 18, 72 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            const string prefix = "Review table";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", length / 2)) : new string('a', length);
            if (zeroAdvance) { word = "a" + word[1..].Replace('a', 'b'); }
            string tail = " " + word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: 500, tailFirstAdvance: zeroAdvance ? (short)0 : null), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            TestAssert.Equal(" ", shows[1].Text);
            TestAssert.Equal(900, (int)shows[1].Font.Font.GetAdvanceWidth(shows[1].Font.Font.MapCodePoint(' ')));
            double size = shows[1].Size;
            TestAssert.True(Math.Abs(shows[1].Y - shows[0].Y) < 0.001d, "The authored separator stays on the prefix row.");
            TestAssert.True(Math.Abs(shows[1].X - shows[0].X - shows[0].Font.MeasureTextPoints(prefix, size)) < 0.04d,
                "The separator starts after the complete prefix advance.");
            var tailRows = shows.Skip(2).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 1, "The separated word must paint below the prefix.");
            if (zeroAdvance) { TestAssert.Equal(0, (int)tailRows[0].Font.Font.GetAdvanceWidth(tailRows[0].Font.Font.MapCodePoint('a'))); }
            double firstPitch = -prefixDescender / 1000d + 0.8d;
            TestAssert.True(Math.Abs(shows[0].Y - tailRows[0].Y - firstPitch * size) < 0.02d,
                "A tail-face separator alone does not replace the visible prefix's descent.");
            foreach (var row in tailRows)
            {
                TestAssert.Equal(900, (int)row.Font.Font.GetAdvanceWidth(row.Font.Font.MapCodePoint('b')));
                TestAssert.True(row.X + row.Font.MeasureTextPoints(row.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d,
                    "Whole and split first words retain visible continuation width.");
            }
            for (int i = 1; i < tailRows.Length; i++)
            {
                TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - size) < 0.02d,
                    "Later rows retain the tail face's metrics.");
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Height includes the mixed first step and every tail row.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsSeparatedTailAfterTerminalPunctuation()
    {
        foreach (char punctuation in new[] { ',', ':', ';' })
        foreach (bool supplementary in new[] { false, true })
        foreach (bool zeroAdvance in new[] { false, true })
        foreach (bool fittingFirstWord in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (int length in new[] { 18, 72 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            string prefix = "Review table" + punctuation;
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", length / 2)) : new string('a', length);
            if (zeroAdvance) { word = "a" + word[1..].Replace('a', 'b'); }
            string tail = " " + (fittingFirstWord ? "i " : string.Empty) + word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: 500, tailFirstAdvance: zeroAdvance ? (short)0 : null), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            TestAssert.Equal(fittingFirstWord ? " i " : " ", shows[1].Text);
            TestAssert.Equal(900, (int)shows[1].Font.Font.GetAdvanceWidth(shows[1].Font.Font.MapCodePoint(' ')));
            double size = shows[1].Size;
            TestAssert.True(Math.Abs(shows[1].Y - shows[0].Y) < 0.001d, "The authored separator stays on the prefix row.");
            TestAssert.True(Math.Abs(shows[1].X - shows[0].X - shows[0].Font.MeasureTextPoints(prefix, size)) < 0.04d,
                "The separator starts after the complete prefix advance.");
            var tailRows = shows.Skip(2).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 1, "The separated word must paint below the prefix.");
            if (zeroAdvance) { TestAssert.Equal(0, (int)tailRows[0].Font.Font.GetAdvanceWidth(tailRows[0].Font.Font.MapCodePoint('a'))); }
            double firstPitch = (fittingFirstWord ? Math.Max(-prefixDescender / 1000d, 0.2d) : -prefixDescender / 1000d) + 0.8d;
            TestAssert.True(Math.Abs(shows[0].Y - tailRows[0].Y - firstPitch * size) < 0.02d,
                "The first step selects descent from visible text; a separator-only row keeps prefix descent.");
            foreach (var row in tailRows)
            {
                TestAssert.Equal(900, (int)row.Font.Font.GetAdvanceWidth(row.Font.Font.MapCodePoint('b')));
                TestAssert.True(row.X + row.Font.MeasureTextPoints(row.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d,
                    "Whole and split first words retain visible continuation width.");
            }
            for (int i = 1; i < tailRows.Length; i++)
            {
                TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - size) < 0.02d,
                    "Later rows retain the tail face's metrics.");
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Height includes the mixed first step and every tail row.");
            foreach ((string guardedPrefix, string guardedTail) in new[]
            {
                ("Review table::", tail), ("Review, table,", tail),
                ("Review ,", tail), (prefix, "  " + word + " ending.")
            })
            {
                DocxParagraph guardedComment = comment with
                {
                    Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1] with { Text = guardedTail }]
                };
                DocxDocument guarded = document with
                {
                    RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                        !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show =>
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Doubled or internal punctuation, a punctuation-only word or doubled separator keeps complete fallback.");
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonRetainsTwoLeadingTailSpaces()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (bool zeroAdvance in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (int length in new[] { 18, 72 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            const string prefix = "Review table";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", length / 2)) : new string('a', length);
            if (zeroAdvance) { word = "a" + word[1..].Replace('a', 'b'); }
            string tail = "  " + word + " ending.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: 500, tailFirstAdvance: zeroAdvance ? (short)0 : null), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            TestAssert.Equal("  ", shows[1].Text);
            TestAssert.Equal(900, (int)shows[1].Font.Font.GetAdvanceWidth(shows[1].Font.Font.MapCodePoint(' ')));
            double size = shows[1].Size;
            TestAssert.True(Math.Abs(shows[1].Y - shows[0].Y) < 0.001d, "The authored separator stays on the prefix row.");
            TestAssert.True(Math.Abs(shows[1].X - shows[0].X - shows[0].Font.MeasureTextPoints(prefix, size)) < 0.04d,
                "The separator starts after the complete prefix advance.");
            var tailRows = shows.Skip(2).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 1, "The separated word must paint below the prefix.");
            if (zeroAdvance) { TestAssert.Equal(0, (int)tailRows[0].Font.Font.GetAdvanceWidth(tailRows[0].Font.Font.MapCodePoint('a'))); }
            double firstPitch = -prefixDescender / 1000d + 0.8d;
            TestAssert.True(Math.Abs(shows[0].Y - tailRows[0].Y - firstPitch * size) < 0.02d,
                "A tail-face separator alone does not replace the visible prefix's descent.");
            foreach (var row in tailRows)
            {
                TestAssert.Equal(900, (int)row.Font.Font.GetAdvanceWidth(row.Font.Font.MapCodePoint('b')));
                TestAssert.True(row.X + row.Font.MeasureTextPoints(row.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d,
                    "Whole and split first words retain visible continuation width.");
            }
            for (int i = 1; i < tailRows.Length; i++)
            {
                TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - size) < 0.02d,
                    "Later rows retain the tail face's metrics.");
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Height includes the mixed first step and every tail row.");
            foreach ((string guardedPrefix, string guardedTail) in new[]
            {
                (prefix, " ".PadRight(64) + word + "   ending."),
                (prefix, "  " + word + "   ending."),
                (prefix + ",", tail)
            })
            {
                DocxParagraph guardedComment = comment with
                {
                    Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1] with { Text = guardedTail }]
                };
                DocxDocument guarded = document with
                {
                    RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                        !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show =>
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three internal spaces or a punctuated prefix retains complete fallback.");
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingClosingRunAndFinalRowMetrics()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (short prefixAscender in new short[] { 500, 1200 })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (int repeat in new[] { 24, 40 })
        foreach (int closingLength in new[] { 1, 12 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            const string prefix = "Review table";
            string tail = "ok. " + string.Join(" ", Enumerable.Repeat(supplementary ? "a\U0001F600" : "aa", repeat));
            string closingText = new('i', closingLength);
            using MemoryStream stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                    <Default Extension="xml" ContentType="application/xml"/>
                    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                    <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                    </Types>
                    """,
                ["_rels/.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                    </Relationships>
                    """,
                ["word/_rels/document.xml.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                    </Relationships>
                    """,
                ["word/document.xml"] = """
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
                    <w:p><w:r><w:rPr><w:rFonts w:ascii="AnchorFace" w:hAnsi="AnchorFace"/></w:rPr><w:t>Anchor</w:t></w:r><w:r><w:commentReference w:id="1"/></w:r></w:p>
                    </w:body></w:document>
                    """,
                ["word/comments.xml"] = $$"""
                    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer" w:initials="RV"><w:p>
                    <w:pPr><w:rPr><w:rFonts w:ascii="AnchorFace" w:hAnsi="AnchorFace"/><w:sz w:val="36"/></w:rPr></w:pPr>
                    <w:r><w:rPr><w:rFonts w:ascii="CommentFace" w:hAnsi="CommentFace"/></w:rPr><w:t>{{prefix}}</w:t></w:r>
                    <w:r><w:rPr><w:rFonts w:ascii="OtherCommentFace" w:hAnsi="OtherCommentFace"/></w:rPr><w:t>{{tail}}</w:t></w:r>
                    <w:r><w:rPr><w:rFonts w:ascii="CommentFace" w:hAnsi="CommentFace"/></w:rPr><w:t>{{closingText}}</w:t></w:r>
                    </w:p></w:comment></w:comments>
                    """
            });
            DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup) with { MarginRightPoints = rightMargin };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: prefixAscender), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + closingText + " ", string.Concat(shows.Select(show => show.Text)));
            var tailRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('a')) == 900).ToArray();
            TestAssert.True(tailRows.Length >= 3, "The probe must include first, pure-tail and final mixed rows.");
            double size = tailRows[0].Size;
            for (int index = 1; index < tailRows.Length; index++)
            {
                double pitch = index == 1 ? Math.Max(-prefixDescender / 1000d, 0.2d) + 0.8d :
                    index == tailRows.Length - 1 ? 0.2d + Math.Max(prefixAscender / 1000d + prefixGap / 1000d, 0.8d) : 1d;
                TestAssert.True(Math.Abs(tailRows[index - 1].Y - tailRows[index].Y - pitch * size) < 0.02d,
                    "First, pure-tail and final mixed steps use their own visible font metrics.");
            }
            var closing = shows.Single(show => show.Text == closingText);
            TestAssert.Equal(700, (int)closing.Font.Font.GetAdvanceWidth(closing.Font.Font.MapCodePoint('i')));
            TestAssert.True(Math.Abs(closing.Y - tailRows[^1].Y) < 0.001d, "The fitting closing run shares the final tail baseline.");
            TestAssert.True(Math.Abs(closing.X - tailRows[^1].X - tailRows[^1].Font.MeasureTextPoints(tailRows[^1].Text, size)) < 0.04d,
                "The closing run follows the final tail advance without an invented separator.");
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The final paragraph mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + tailRows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Balloon height includes all three row phases.");
            DocxRelatedStory story = document.RelatedStories.Single();
            DocxParagraph comment = ((DocxParagraphElement)story.BodyElements.Single()).Paragraph;
            TestAssert.Equal(" ", comment.ParagraphMarkRun!.Text);
            TestAssert.Equal("AnchorFace", comment.ParagraphMarkRun.FontFamily);
            foreach (DocxTextRun[] guardedRuns in new[]
            {
                new[] { comment.Runs[0], comment.Runs[1], comment.Runs[2] with { Text = new string('a', 200) } },
                [comment.Runs[0], comment.Runs[1], comment.Runs[2] with { Bold = true }],
                [comment.Runs[0], comment.Runs[1] with { Text = " " + tail }, comment.Runs[2]],
                [comment.Runs[0], comment.Runs[1], comment.Runs[2] with { FontFamily = "OtherCommentFace",
                    Fonts = comment.Runs[2].Fonts with { Ascii = "OtherCommentFace", HighAnsi = "OtherCommentFace" } }],
                [comment.Runs[0], comment.Runs[1], comment.Runs[2], comment.Runs[2] with { Text = "end" }],
                [comment.Runs[0] with { Text = "Review table," }, comment.Runs[1], comment.Runs[2]]
            })
            {
                DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(comment with { Runs = guardedRuns })] }] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    $"Guard retains complete fallback: runs={guardedRuns.Length}, closingLength={guardedRuns[2].Text.Length}, bold={guardedRuns[2].Bold}, closingFace={guardedRuns[2].FontFamily}, leadingTail={guardedRuns[1].Text.StartsWith(' ')}, prefixComma={guardedRuns[0].Text.EndsWith(',')}.");
            }
            foreach (short advance in new short[] { 0, 32700 })
            {
                DocxParagraph guardedComment = comment with
                {
                    Runs = [comment.Runs[0], comment.Runs[1] with { Text = "ok. " + new string('a', 30) + " end." }, comment.Runs[2]]
                };
                DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }] };
                var numericRenderer = new DocxRenderer(new BalloonTypefaceFontResolver(tailFirstAdvance: advance), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var guardedBalloon = numericRenderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Zero-width and overwide tail words retain complete fallback in the closing-run branch.");
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonMovesPreservedClosingRunToOwnRow()
    {
        var first = DocxTests.CreateDocxLayoutParagraph("end.", 12d, 14d) with
        {
            LineSpacingPoints = null,
            Runs = [new DocxTextRun("end.", 12d, null, false, false, false, null, "CommentFace")],
            ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
        };
        string prefix = new string('a', 80) + " finish.";
        string closing = " " + new string('m', 15) + ".";
        var terminal = first with
        {
            Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
            ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
        };
        var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
        {
            Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
        };
        var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
            [new DocxParagraphElement(first), new DocxParagraphElement(first), new DocxParagraphElement(terminal)], [], [], null)
        {
            CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
        };
        var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
            [], [], [], [new DocxParagraphElement(anchor)], [], [])
        {
            RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodySpaceAdvance: 123),
            OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var balloon = renderer.InspectMarkupBalloons(document).Single();
        var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
            .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal("end. end. " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
        var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
        var last = rows[^1].ToArray();
        var preceding = rows[^2].ToArray();
        TestAssert.Equal(closing.TrimStart(' '), last[0].Text);
        TestAssert.Equal(" ", last[1].Text);
        TestAssert.Equal(2, last.Length);
        TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
            "A preserved closing separator keeps the preceding whole word on its original row.");
        TestAssert.Equal(" ", preceding[^1].Text);
        TestAssert.Equal(last[0].Font, preceding[^1].Font);
        double size = last[0].Size;
        var previousFont = preceding[0].Font.Font;
        var closingFont = last[0].Font.Font;
        double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
            (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
        TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
            "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
        TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
            Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
            "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
        TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
            "Height includes the independent closing-only transition.");
    }

    public static void DocxWordCompatibleMixedBalloonMovesPreservedSecondClosingRunToOwnRow()
    {
        foreach (bool wrappedLeading in new[] { false, true })
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = wrappedLeading ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " " + new string('m', 15) + ".";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(first), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(leadingText + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(closing.TrimStart(' '), last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonMovesPreservedFourthClosingRunToOwnRow()
    {
        foreach (int wrapMask in Enumerable.Range(0, 8))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 3).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " " + new string('m', 15) + ".";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 5, "Three preceding bodies and two terminal rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(closing.TrimStart(' '), last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonMovesPreservedFifthClosingRunToOwnRow()
    {
        foreach (int wrapMask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " " + new string('m', 15) + ".";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 7, "Four preceding bodies and terminal source rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(closing.TrimStart(' '), last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
            foreach (DocxParagraph rejected in new[]
            {
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 120) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm mmmmmmm mmmmmmm fourth." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + closing }] },
                terminal with { Runs = [terminal.Runs[0] with { Text = prefix + " " }, terminal.Runs[1] with { Text = closing.TrimStart(' ') }] },
                terminal with { Runs = [.. terminal.Runs, terminal.Runs[0] with { Text = " extra." }] },
                terminal with { Runs = [terminal.Runs[0] with { Bold = true }, terminal.Runs[1]] },
                terminal with { Spacing = terminal.Spacing with { AfterLinesValue = "400" } }
            })
            {
                var rejectedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(rejected)] };
                var rejectedDocument = document with { RelatedStories = [rejectedStory] };
                var rejectedBalloon = renderer.InspectMarkupBalloons(rejectedDocument).Single();
                var rejectedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(rejectedDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= rejectedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(rejectedShows.Length > 0 && rejectedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Overwide/three-word/doubled closing separators, trailing prefixes, extra runs, decoration and unqualified spacing retain fallback.");
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonMovesPreservedTwoWordFifthClosingRunToOwnRow()
    {
        foreach (int wrapMask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " " + new string('m', 15) + " m.";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 7, "Four preceding bodies and terminal source rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(closing.TrimStart(' '), last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
            foreach (DocxParagraph rejected in new[]
            {
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 120) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a " + new string('m', 120) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmmmmmmmmmmmmmmmmmmmmmmm mmmmmmmmmmmmmmmmmmmmmmmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + closing }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm  mmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm mmmmmmm mmmmmmm fourth." }] },
                terminal with { Runs = [terminal.Runs[0] with { Text = prefix + " " }, terminal.Runs[1] with { Text = closing.TrimStart(' ') }] },
                terminal with { Runs = [.. terminal.Runs, terminal.Runs[0] with { Text = " extra." }] },
                terminal with { Runs = [terminal.Runs[0] with { Bold = true }, terminal.Runs[1]] },
                terminal with { Spacing = terminal.Spacing with { AfterLinesValue = "400" } }
            })
            {
                var rejectedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(rejected)] };
                var rejectedDocument = document with { RelatedStories = [rejectedStory] };
                var rejectedBalloon = renderer.InspectMarkupBalloons(rejectedDocument).Single();
                var rejectedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(rejectedDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= rejectedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(rejectedShows.Length > 0 && rejectedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Overwide two-word/doubled closing separators, trailing prefixes, extra runs, decoration and unqualified spacing retain fallback: " + rejected.Runs[1].Text);
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonMovesPreservedThreeWordFifthClosingRunToOwnRow()
    {
        foreach (int wrapMask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " " + new string('m', 15) + " m m.";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 7, "Four preceding bodies and terminal source rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(closing.TrimStart(' '), last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
            foreach (DocxParagraph rejected in new[]
            {
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 120) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a " + new string('m', 120) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmmmmmmmmmmmmmmmmmmmmmmm mmmmmmmmmmmmmmmmmmmmmmmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + closing }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm  mmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm mmmmmmm mmmmmmm fourth." }] },
                terminal with { Runs = [terminal.Runs[0] with { Text = prefix + " " }, terminal.Runs[1] with { Text = closing.TrimStart(' ') }] },
                terminal with { Runs = [.. terminal.Runs, terminal.Runs[0] with { Text = " extra." }] },
                terminal with { Runs = [terminal.Runs[0] with { Bold = true }, terminal.Runs[1]] },
                terminal with { Spacing = terminal.Spacing with { AfterLinesValue = "400" } }
            })
            {
                var rejectedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(rejected)] };
                var rejectedDocument = document with { RelatedStories = [rejectedStory] };
                var rejectedBalloon = renderer.InspectMarkupBalloons(rejectedDocument).Single();
                var rejectedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(rejectedDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= rejectedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(rejectedShows.Length > 0 && rejectedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Overwide two-word/doubled closing separators, trailing prefixes, extra runs, decoration and unqualified spacing retain fallback: " + rejected.Runs[1].Text);
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingFirstWordOfFifthClosingRun()
    {
        foreach (int wrapMask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " a " + new string('m', 15) + ".";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 7, "Four preceding bodies and terminal source rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(new string('m', 15) + ".", last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" a ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double mixedPitch = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                Math.Max((previousFont.Hhea.HorizontalAscender + previousFont.Hhea.HorizontalLineGap) / (double)previousFont.UnitsPerEm,
                    (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm)) * size;
            TestAssert.True(Math.Abs(rows[^3].Key - rows[^2].Key - mixedPitch) < 0.02d,
                "The last first-body row includes the fitting closing face's ascent/gap.");
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
            foreach (DocxParagraph rejected in new[]
            {
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 120) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a " + new string('m', 120) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmmmmmmmmmmmmmmmmmmmmmmm mmmmmmmmmmmmmmmmmmmmmmmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + closing }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm  mmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm mmmmmmm mmmmmmm fourth." }] },
                terminal with { Runs = [terminal.Runs[0] with { Text = prefix + " " }, terminal.Runs[1] with { Text = closing.TrimStart(' ') }] },
                terminal with { Runs = [.. terminal.Runs, terminal.Runs[0] with { Text = " extra." }] },
                terminal with { Runs = [terminal.Runs[0] with { Bold = true }, terminal.Runs[1]] },
                terminal with { Spacing = terminal.Spacing with { AfterLinesValue = "400" } }
            })
            {
                var rejectedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(rejected)] };
                var rejectedDocument = document with { RelatedStories = [rejectedStory] };
                var rejectedBalloon = renderer.InspectMarkupBalloons(rejectedDocument).Single();
                var rejectedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(rejectedDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= rejectedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(rejectedShows.Length > 0 && rejectedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Overwide two-word/doubled closing separators, trailing prefixes, extra runs, decoration and unqualified spacing retain fallback: " + rejected.Runs[1].Text);
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingFirstWordOfThreeWordFifthClosingRun()
    {
        foreach (int wrapMask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " a " + new string('m', 15) + ". third.";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 7, "Four preceding bodies and terminal source rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(new string('m', 15) + ". third.", last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" a ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double mixedPitch = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                Math.Max((previousFont.Hhea.HorizontalAscender + previousFont.Hhea.HorizontalLineGap) / (double)previousFont.UnitsPerEm,
                    (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm)) * size;
            TestAssert.True(Math.Abs(rows[^3].Key - rows[^2].Key - mixedPitch) < 0.02d,
                "The last first-body row includes the fitting closing face's ascent/gap.");
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
            foreach (DocxParagraph rejected in new[]
            {
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 120) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a " + new string('m', 120) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmmmmmmmmmmmmmmmmmmmmmmm mmmmmmmmmmmmmmmmmmmmmmmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + closing }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a b c " + new string('m', 15) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = closing + " fourth." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 20) + " abc def." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = closing.Replace(". third.", ".  third.", StringComparison.Ordinal) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm  mmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm mmmmmmm mmmmmmm fourth." }] },
                terminal with { Runs = [terminal.Runs[0] with { Text = prefix + " " }, terminal.Runs[1] with { Text = closing.TrimStart(' ') }] },
                terminal with { Runs = [.. terminal.Runs, terminal.Runs[0] with { Text = " extra." }] },
                terminal with { Runs = [terminal.Runs[0] with { Bold = true }, terminal.Runs[1]] },
                terminal with { Spacing = terminal.Spacing with { AfterLinesValue = "400" } }
            })
            {
                var rejectedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(rejected)] };
                var rejectedDocument = document with { RelatedStories = [rejectedStory] };
                var rejectedBalloon = renderer.InspectMarkupBalloons(rejectedDocument).Single();
                var rejectedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(rejectedDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= rejectedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(rejectedShows.Length > 0 && rejectedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Unqualified three-word boundaries retain fallback: " + rejected.Runs[1].Text);
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingFirstTwoWordsOfThreeWordFifthClosingRun()
    {
        foreach (int wrapMask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            string closing = " a b " + new string('m', 15) + ".";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 7, "Four preceding bodies and terminal source rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(new string('m', 15) + ".", last[0].Text);
            TestAssert.Equal(" ", last[1].Text);
            TestAssert.Equal(2, last.Length);
            TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                "A preserved closing separator keeps the preceding whole word on its original row.");
            TestAssert.Equal(" a b ", preceding[^1].Text);
            TestAssert.Equal(last[0].Font, preceding[^1].Font);
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[0].Font.Font;
            double mixedPitch = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                Math.Max((previousFont.Hhea.HorizontalAscender + previousFont.Hhea.HorizontalLineGap) / (double)previousFont.UnitsPerEm,
                    (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm)) * size;
            TestAssert.True(Math.Abs(rows[^3].Key - rows[^2].Key - mixedPitch) < 0.02d,
                "The last first-body row includes the fitting closing face's ascent/gap.");
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "A closing-only row uses its own ascent/gap and the preceding first-body descent.");
            TestAssert.True(Math.Abs(last[0].X - preceding[0].X) < 0.001d &&
                Math.Abs(last[1].X - last[0].X - last[0].Font.MeasureTextPoints(last[0].Text, size)) < 0.04d,
                "The closing-only row resets to the continuation inset and retains its nominal mark advance.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height includes the independent closing-only transition.");
            foreach (DocxParagraph rejected in new[]
            {
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 120) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a " + new string('m', 120) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmmmmmmmmmmmmmmmmmmmmmmm mmmmmmmmmmmmmmmmmmmmmmmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + closing }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " a b c " + new string('m', 15) + "." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = closing + " fourth." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " " + new string('m', 20) + " abc def." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = closing.Replace(" a b ", " a b  ", StringComparison.Ordinal) }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm  mmmmmmm." }] },
                terminal with { Runs = [terminal.Runs[0], terminal.Runs[1] with { Text = " mmmmmmm mmmmmmm mmmmmmm fourth." }] },
                terminal with { Runs = [terminal.Runs[0] with { Text = prefix + " " }, terminal.Runs[1] with { Text = closing.TrimStart(' ') }] },
                terminal with { Runs = [.. terminal.Runs, terminal.Runs[0] with { Text = " extra." }] },
                terminal with { Runs = [terminal.Runs[0] with { Bold = true }, terminal.Runs[1]] },
                terminal with { Spacing = terminal.Spacing with { AfterLinesValue = "400" } }
            })
            {
                var rejectedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(rejected)] };
                var rejectedDocument = document with { RelatedStories = [rejectedStory] };
                var rejectedBalloon = renderer.InspectMarkupBalloons(rejectedDocument).Single();
                var rejectedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(rejectedDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= rejectedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(rejectedShows.Length > 0 && rejectedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Unqualified three-word boundaries retain fallback: " + rejected.Runs[1].Text);
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsMultiwordFourthClosingReflow()
    {
        foreach (int wrapMask in Enumerable.Range(0, 8))
        foreach (bool supplementary in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string leadingText = (wrapMask & 1) != 0 ? new string('m', 72) + " ending." : "end.";
            var first = DocxTests.CreateDocxLayoutParagraph(leadingText, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(leadingText, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            var leading = Enumerable.Range(0, 3).Select(index =>
            {
                string text = (wrapMask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return first with { Runs = [first.Runs[0] with { Text = text, FontFamily = index == 1 ? "OtherCommentFace" : "CommentFace" }] };
            }).ToArray();
            string prefix = new string('a', 80) + (supplementary ? "\U0001F600" : "") + " finish.";
            const string closing = " mmmmmmm mmmmmmm.";
            var terminal = first with
            {
                Runs = [first.Runs[0] with { Text = prefix }, first.Runs[0] with { Text = closing, FontFamily = "OtherCommentFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            };
            if (canonicalSpacing)
            {
                first = first with { Spacing = first.Spacing with
                {
                    BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                    BeforeValue = "120", AfterValue = "240"
                } };
                leading = leading.Select(paragraph => paragraph with { Spacing = first.Spacing }).ToArray();
                terminal = terminal with { Spacing = first.Spacing };
            }
            var anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            var story = new DocxRelatedStory(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [.. leading.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)), new DocxParagraphElement(terminal)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            var document = new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length > 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Join(" ", leading.Select(paragraph => paragraph.Runs[0].Text)) + " " + prefix + closing + " ", string.Concat(shows.Select(show => show.Text)));
            var rows = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(rows.Length >= 5, "Three preceding bodies and wrapped terminal rows remain distinct.");
            var last = rows[^1].ToArray();
            var preceding = rows[^2].ToArray();
            TestAssert.Equal(3, last.Length);
            TestAssert.True(last[0].Text.EndsWith("finish.", StringComparison.Ordinal), "Existing mixed reflow keeps the complete preceding word.");
            TestAssert.Equal(closing, last[1].Text);
            TestAssert.Equal(" ", last[2].Text);
            TestAssert.Equal(700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[2].Font.Font.GetAdvanceWidth(last[2].Font.Font.MapCodePoint(' ')));
            double size = last[0].Size;
            var previousFont = preceding[0].Font.Font;
            var closingFont = last[1].Font.Font;
            double transition = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm +
                Math.Max((previousFont.Hhea.HorizontalAscender + previousFont.Hhea.HorizontalLineGap) / (double)previousFont.UnitsPerEm,
                    (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm)) * size;
            TestAssert.True(Math.Abs(preceding[0].Y - last[0].Y - transition) < 0.02d,
                "Existing multiword mixed reflow retains its larger body ascent/gap and preceding first-body descent.");
            for (int index = 0; index < 2; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, size)) < 0.04d,
                    "Existing multiword fourth-row body parts and mark retain their own advances.");
            }
            TestAssert.True(Math.Abs(last[2].Size - last[0].Size) < 0.001d,
                "Existing multiword closing reflow retains its nominal mark.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Existing multiword reflow height and emission retain their shared mixed-row transition.");
        }
    }

    public static void DocxWordCompatibleMixedBalloonRetainsLeadingAndInternalTailSpaceSpans()
    {
        foreach (bool doubleInternalSeparator in new[] { false, true })
        foreach (bool supplementary in new[] { false, true })
        foreach (bool zeroAdvance in new[] { false, true })
        foreach (int spaceCount in new[] { 3, 4, 8, 12, 256, 1024 })
        foreach (bool fittingFirstWord in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (int length in new[] { 18, 72 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            if (doubleInternalSeparator && fittingFirstWord) { continue; }
            const string prefix = "Review table";
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", length / 2)) : new string('a', length);
            if (zeroAdvance) { word = "a" + word[1..].Replace('a', 'b'); }
            string spaces = " ".PadRight(spaceCount);
            string tail = spaces + (fittingFirstWord ? "i " : string.Empty) + word +
                (doubleInternalSeparator ? "  ending." : " ending.");
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                    new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(comment)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: 450, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: 500, tailFirstAdvance: zeroAdvance ? (short)0 : null, tailSpaceAdvance: 123), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
            bool visibleTailOnFirstRow = fittingFirstWord && spaceCount < 256;
            TestAssert.Equal(visibleTailOnFirstRow ? spaces + "i " : spaces, shows[1].Text);
            TestAssert.Equal(123, (int)shows[1].Font.Font.GetAdvanceWidth(shows[1].Font.Font.MapCodePoint(' ')));
            double size = shows[1].Size;
            TestAssert.True(Math.Abs(shows[1].Y - shows[0].Y) < 0.001d, "The authored separator stays on the prefix row.");
            TestAssert.True(Math.Abs(shows[1].X - shows[0].X - shows[0].Font.MeasureTextPoints(prefix, size)) < 0.04d,
                "The separator starts after the complete prefix advance.");
            var tailRows = shows.Skip(2).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
            TestAssert.True(tailRows.Length >= 1, "The separated word must paint below the prefix.");
            if (zeroAdvance) { TestAssert.Equal(0, (int)tailRows[0].Font.Font.GetAdvanceWidth(tailRows[0].Font.Font.MapCodePoint('a'))); }
            if (spaceCount >= 256)
            {
                TestAssert.True(shows[1].X + shows[1].Font.MeasureTextPoints(spaces, size) > balloon.X + balloon.Width,
                    "An overflowing authored space span stays complete beside its prefix.");
                TestAssert.True(double.IsFinite(balloon.Height), "Overflowing whitespace retains finite layout.");
            }
            double firstPitch = (visibleTailOnFirstRow ? Math.Max(-prefixDescender / 1000d, 0.2d) : -prefixDescender / 1000d) + 0.8d;
            TestAssert.True(Math.Abs(shows[0].Y - tailRows[0].Y - firstPitch * size) < 0.02d,
                "A tail-face separator alone does not replace the visible prefix's descent.");
            foreach (var row in tailRows)
            {
                TestAssert.Equal(900, (int)row.Font.Font.GetAdvanceWidth(row.Font.Font.MapCodePoint('b')));
                TestAssert.True(row.X + row.Font.MeasureTextPoints(row.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d,
                    "Whole and split first words retain visible continuation width.");
            }
            for (int i = 1; i < tailRows.Length; i++)
            {
                TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - size) < 0.02d,
                    "Later rows retain the tail face's metrics.");
            }
            TestAssert.Equal(" ", shows[^1].Text);
            TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - tailRows[^1].Y)) < 0.02d,
                "Height includes the mixed first step and every tail row.");
            foreach ((string guardedPrefix, string guardedTail) in new[]
            {
                (prefix, " ".PadRight(256) + word + "   ending."),
                (prefix, spaces + word + "   ending."),
                (prefix, " " + word + "  ending."),
                (prefix, spaces + "ok. " + word + "  ending."),
                (prefix + ",", tail), (prefix, "\t" + word + " ending.")
            })
            {
                DocxParagraph guardedComment = comment with
                {
                    Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1] with { Text = guardedTail }]
                };
                DocxDocument guarded = document with
                {
                    RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                        !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show =>
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three internal spaces, a single leading separator, extra tail words, punctuated prefix or tab retains complete fallback.");
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonSplitsOverwidePrefixWordInItsSourceFace()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (int prefixLength in new[] { 80, 120 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            string prefix = new string('a', prefixLength);
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            foreach (string tail in new[] { "ok. end.", "ok. " + word + " ending.", word + " ending." })
            {
                DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
                {
                    Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
                };
                DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
                {
                    Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                        new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
                DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                    [new DocxParagraphElement(comment)], [], [], null)
                {
                    CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
                };
                DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                    [], [], [], [new DocxParagraphElement(anchor)], [], [])
                {
                    RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
                };
                var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: 1800,
                    bodyDescender: -350, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var balloon = renderer.InspectMarkupBalloons(document).Single();
                var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                var prefixRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 700).ToArray();
                var tailRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('a')) == 900).ToArray();
                TestAssert.True(prefixRows.Length >= 3, "Ordinary prefix words must wrap in their prepared face.");
                TestAssert.True(tailRows.Length >= 1, "The mixed row must retain its prepared tail face.");
                TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
                TestAssert.True(prefixRows.All(row => row.Text.All(character => character == 'a')), "Prefix hard breaks must not invent separators.");
                double size = tailRows[0].Size;
                double prefixPitch = 0.85d + prefixGap / 1000d;
                double incomingPitch = 0.35d + Math.Max(0.5d + prefixGap / 1000d, 0.8d);
                for (int i = 1; i < prefixRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(prefixRows[i - 1].Y - prefixRows[i].Y -
                        (i == prefixRows.Length - 1 ? incomingPitch : prefixPitch) * size) < 0.02d,
                        "Prefix-only rows and the incoming mixed row use their distinct metrics.");
                }
                TestAssert.True(Math.Abs(prefixRows[^1].Y - tailRows[0].Y) < 0.001d, "The tail joins the final prefix row.");
                double tailX = prefixRows[^1].X + prefixRows[^1].Font.MeasureTextPoints(prefixRows[^1].Text, size);
                TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "Joined tail position uses the final prefix advance.");
                for (int i = 1; i < tailRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - (i == 1 ? 1.15d : 1d) * size) < 0.02d,
                        "The outgoing mixed step uses both descents; later rows use the tail face alone.");
                }
                TestAssert.True(prefixRows.Concat(tailRows).All(show =>
                    show.X + show.Font.MeasureTextPoints(show.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d),
                    "Visible words must stay inside the balloon.");
                TestAssert.Equal(" ", shows[^1].Text);
                TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
                TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
                TestAssert.True(Math.Abs(balloon.Height - (12.61d + prefixRows[0].Y - tailRows[^1].Y)) < 0.02d,
                    "Height includes prefix-only, incoming mixed, outgoing mixed and tail-only steps.");

                foreach (short firstAdvance in new short[] { 0, 32700 })
                {
                    DocxParagraph guardedComment = comment with
                    {
                        Runs = [comment.Runs[0], comment.Runs[1] with { Text = "a" + new string('b', 72) + " ending." }]
                    };
                    DocxDocument guarded = document with
                    {
                        RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                    };
                    var guardedRenderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap,
                        bodySpaceAdvance: 1800, bodyDescender: -350, bodyAscender: 500, tailFirstAdvance: firstAdvance),
                        OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var guardedBalloon = guardedRenderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(guardedRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                            !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "A zero-width or oversized first scalar retains the complete fallback path.");
                }
                if (tail == "ok. end.")
                {
                    foreach ((string guardedPrefix, short? firstAdvance) in new (string, short?)[]
                    {
                        ("a" + new string('b', prefixLength), 0),
                        ("a" + new string('b', prefixLength), 32700),
                        (prefix + ";;", null),
                        (prefix + "\U0001F600", null)
                    })
                    {
                        DocxParagraph guardedComment = comment with
                        {
                            Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1]]
                        };
                        DocxDocument guarded = document with
                        {
                            RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                        };
                        var guardedRenderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap,
                            bodySpaceAdvance: 1800, bodyDescender: -350, bodyAscender: 500, bodyFirstAdvance: firstAdvance),
                            OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                        var guardedBalloon = guardedRenderer.InspectMarkupBalloons(guarded).Single();
                        var guardedShows = ReadEmbeddedGlyphTextShows(guardedRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                            .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                                !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                        TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show =>
                            show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                            "A zero-width or oversized first prefix glyph, punctuation or non-ASCII prefix retains the complete fallback path.");
                    }
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonSplitsOverwideTerminalPunctuationPrefixWord()
    {
        foreach (char punctuation in new[] { ',', ':', ';' })
        foreach (bool supplementary in new[] { false, true })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (int prefixLength in new[] { 80, 120 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            string prefix = new string('a', prefixLength) + punctuation;
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            foreach (string tail in new[] { "ok. end.", "ok. " + word + " ending.", word + " ending." })
            {
                DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
                {
                    Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
                };
                DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
                {
                    Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                        new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
                DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                    [new DocxParagraphElement(comment)], [], [], null)
                {
                    CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
                };
                DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                    [], [], [], [new DocxParagraphElement(anchor)], [], [])
                {
                    RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
                };
                var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: 1800,
                    bodyDescender: -350, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var balloon = renderer.InspectMarkupBalloons(document).Single();
                var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                var prefixRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 700).ToArray();
                var tailRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('a')) == 900).ToArray();
                TestAssert.True(prefixRows.Length >= 3, "Ordinary prefix words must wrap in their prepared face.");
                TestAssert.True(tailRows.Length >= 1, "The mixed row must retain its prepared tail face.");
                TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
                TestAssert.True(prefixRows.All(row => row.Text.All(character => character == 'a' || character == punctuation)), "Prefix hard breaks must not invent separators.");
                double size = tailRows[0].Size;
                double prefixPitch = 0.85d + prefixGap / 1000d;
                double incomingPitch = 0.35d + Math.Max(0.5d + prefixGap / 1000d, 0.8d);
                for (int i = 1; i < prefixRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(prefixRows[i - 1].Y - prefixRows[i].Y -
                        (i == prefixRows.Length - 1 ? incomingPitch : prefixPitch) * size) < 0.02d,
                        "Prefix-only rows and the incoming mixed row use their distinct metrics.");
                }
                TestAssert.True(Math.Abs(prefixRows[^1].Y - tailRows[0].Y) < 0.001d, "The tail joins the final prefix row.");
                double tailX = prefixRows[^1].X + prefixRows[^1].Font.MeasureTextPoints(prefixRows[^1].Text, size);
                TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "Joined tail position uses the final prefix advance.");
                for (int i = 1; i < tailRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - (i == 1 ? 1.15d : 1d) * size) < 0.02d,
                        "The outgoing mixed step uses both descents; later rows use the tail face alone.");
                }
                TestAssert.True(prefixRows.Concat(tailRows).All(show =>
                    show.X + show.Font.MeasureTextPoints(show.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d),
                    "Visible words must stay inside the balloon.");
                TestAssert.Equal(" ", shows[^1].Text);
                TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
                TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
                TestAssert.True(Math.Abs(balloon.Height - (12.61d + prefixRows[0].Y - tailRows[^1].Y)) < 0.02d,
                    "Height includes prefix-only, incoming mixed, outgoing mixed and tail-only steps.");

                foreach (short firstAdvance in new short[] { 0, 32700 })
                {
                    DocxParagraph guardedComment = comment with
                    {
                        Runs = [comment.Runs[0], comment.Runs[1] with { Text = "a" + new string('b', 72) + " ending." }]
                    };
                    DocxDocument guarded = document with
                    {
                        RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                    };
                    var guardedRenderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap,
                        bodySpaceAdvance: 1800, bodyDescender: -350, bodyAscender: 500, tailFirstAdvance: firstAdvance),
                        OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var guardedBalloon = guardedRenderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(guardedRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                            !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "A zero-width or oversized first scalar retains the complete fallback path.");
                }
                if (tail == "ok. end.")
                {
                    foreach ((string guardedPrefix, short? firstAdvance) in new (string, short?)[]
                    {
                        ("a" + new string('b', prefixLength) + punctuation, 0),
                        ("a" + new string('b', prefixLength) + punctuation, 32700),
                        (prefix + ",", null),
                        (prefix + "\U0001F600", null)
                    })
                    {
                        DocxParagraph guardedComment = comment with
                        {
                            Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1]]
                        };
                        DocxDocument guarded = document with
                        {
                            RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                        };
                        var guardedRenderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap,
                            bodySpaceAdvance: 1800, bodyDescender: -350, bodyAscender: 500, bodyFirstAdvance: firstAdvance),
                            OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                        var guardedBalloon = guardedRenderer.InspectMarkupBalloons(guarded).Single();
                        var guardedShows = ReadEmbeddedGlyphTextShows(guardedRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                            .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                                !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                        TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show =>
                            show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                            "A zero-width or oversized first prefix glyph, punctuation or non-ASCII prefix retains the complete fallback path.");
                    }
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonSplitsFirstTailWordAfterWrappedPrefix()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (short prefixSpace in new short[] { 700, 1800 })
        foreach (int prefixCount in new[] { 12, 13 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            string prefix = string.Join(" ", Enumerable.Repeat("Review table", prefixCount));
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            string boundaryWord = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 12)) + "a" : new string('a', 25);
            foreach (string tail in new[] { word + " ending.", boundaryWord + " ending.", boundaryWord + " " + new string('b', 72) + " ending." })
            {
                DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
                {
                    Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
                };
                DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
                {
                    Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                        new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
                DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                    [new DocxParagraphElement(comment)], [], [], null)
                {
                    CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
                };
                DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                    [], [], [], [new DocxParagraphElement(anchor)], [], [])
                {
                    RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
                };
                var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: prefixSpace,
                    bodyDescender: -350, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var balloon = renderer.InspectMarkupBalloons(document).Single();
                var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                var prefixRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 700).ToArray();
                var tailRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('a')) == 900).ToArray();
                TestAssert.True(prefixRows.Length >= 4, "Ordinary prefix words must wrap in their prepared face.");
                TestAssert.True(tailRows.Length >= 2, "The mixed row must retain its prepared tail face.");
                TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
                TestAssert.Equal("table", prefixRows[^1].Text);
                if (prefixSpace == 1800) { TestAssert.Equal("Review table Review ", prefixRows[0].Text); }
                double size = tailRows[0].Size;
                double prefixPitch = 0.85d + prefixGap / 1000d;
                double incomingPitch = 0.35d + Math.Max(0.5d + prefixGap / 1000d, 0.8d);
                for (int i = 1; i < prefixRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(prefixRows[i - 1].Y - prefixRows[i].Y -
                        (i == prefixRows.Length - 1 ? incomingPitch : prefixPitch) * size) < 0.02d,
                        "Prefix-only rows and the incoming mixed row use their distinct metrics.");
                }
                TestAssert.True(Math.Abs(prefixRows[^1].Y - tailRows[0].Y) < 0.001d, "The tail joins the final prefix row.");
                double tailX = prefixRows[^1].X + prefixRows[^1].Font.MeasureTextPoints(prefixRows[^1].Text, size);
                TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "Joined tail position uses the final prefix advance.");
                TestAssert.True(!tailRows[0].Text.EndsWith(' ') && tailRows[0].Text.Length < tail.Split(' ')[0].Length,
                    "The joined first tail word splits without inventing a separator.");
                for (int i = 1; i < tailRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - (i == 1 ? 1.15d : 1d) * size) < 0.02d,
                        "The outgoing mixed step uses both descents; later rows use the tail face alone.");
                }
                TestAssert.True(prefixRows.Concat(tailRows).All(show =>
                    show.X + show.Font.MeasureTextPoints(show.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d),
                    "Visible words must stay inside the balloon.");
                TestAssert.Equal(" ", shows[^1].Text);
                TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
                TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
                TestAssert.True(Math.Abs(balloon.Height - (12.61d + prefixRows[0].Y - tailRows[^1].Y)) < 0.02d,
                    "Height includes prefix-only, incoming mixed, outgoing mixed and tail-only steps.");

                foreach ((string guardedPrefix, string guardedTail) in new[]
                {
                    (word + ";;", "ok. ending."), (prefix.Replace("table", "table,"), tail)
                })
                {
                    DocxParagraph guardedComment = comment with
                    {
                        Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1] with { Text = guardedTail }]
                    };
                    DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }] };
                    var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                            !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Overwide prefix words and punctuation retain the complete fallback path.");
                }
                foreach (short firstAdvance in new short[] { 0, 32700 })
                {
                    DocxParagraph guardedComment = comment with
                    {
                        Runs = [comment.Runs[0], comment.Runs[1] with { Text = "a" + new string('b', 72) + " ending." }]
                    };
                    DocxDocument guarded = document with
                    {
                        RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }]
                    };
                    var guardedRenderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap,
                        bodySpaceAdvance: prefixSpace, bodyDescender: -350, bodyAscender: 500, tailFirstAdvance: firstAdvance),
                        OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var guardedBalloon = guardedRenderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(guardedRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                            !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "A zero-width or oversized first scalar retains the complete fallback path.");
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsOrdinaryPrefixWordsBeforeJoiningTail()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (short prefixSpace in new short[] { 700, 1800 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            string prefix = string.Join(" ", Enumerable.Repeat("Review table", 12));
            string word = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 36)) : new string('a', 72);
            string boundaryWord = supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 7)) : new string('a', 14);
            foreach (string tail in new[] { "ok. end.", "ok. " + word + " ending.", boundaryWord + " " + new string('b', 8) + " ending." })
            {
                DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
                {
                    Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
                };
                DocxParagraph comment = DocxTests.CreateDocxLayoutParagraph(prefix + tail, 12d, 14d) with
                {
                    Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace"),
                        new DocxTextRun(tail, 12d, null, false, false, false, null, "OtherCommentFace")],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
                DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                    [new DocxParagraphElement(comment)], [], [], null)
                {
                    CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
                };
                DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                    [], [], [], [new DocxParagraphElement(anchor)], [], [])
                {
                    RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
                };
                var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: prefixSpace,
                    bodyDescender: -350, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var balloon = renderer.InspectMarkupBalloons(document).Single();
                var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                var prefixRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 700).ToArray();
                var tailRows = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) &&
                    show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('a')) == 900).ToArray();
                TestAssert.True(prefixRows.Length >= 4, "Ordinary prefix words must wrap in their prepared face.");
                TestAssert.True(tailRows.Length >= 1, "The mixed row must retain its prepared tail face.");
                TestAssert.Equal(prefix + tail + " ", string.Concat(shows.Select(show => show.Text)));
                TestAssert.Equal("table", prefixRows[^1].Text);
                if (prefixSpace == 1800) { TestAssert.Equal("Review table Review ", prefixRows[0].Text); }
                double size = tailRows[0].Size;
                double prefixPitch = 0.85d + prefixGap / 1000d;
                double incomingPitch = 0.35d + Math.Max(0.5d + prefixGap / 1000d, 0.8d);
                for (int i = 1; i < prefixRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(prefixRows[i - 1].Y - prefixRows[i].Y -
                        (i == prefixRows.Length - 1 ? incomingPitch : prefixPitch) * size) < 0.02d,
                        "Prefix-only rows and the incoming mixed row use their distinct metrics.");
                }
                TestAssert.True(Math.Abs(prefixRows[^1].Y - tailRows[0].Y) < 0.001d, "The tail joins the final prefix row.");
                double tailX = prefixRows[^1].X + prefixRows[^1].Font.MeasureTextPoints(prefixRows[^1].Text, size);
                TestAssert.True(Math.Abs(tailRows[0].X - tailX) < 0.04d, "Joined tail position uses the final prefix advance.");
                if (tail.StartsWith(boundaryWord, StringComparison.Ordinal))
                {
                    TestAssert.Equal(boundaryWord + " " + new string('b', 8) + " ", tailRows[0].Text);
                }
                if (tail == "ok. end.") { TestAssert.Equal(1, tailRows.Length); }
                for (int i = 1; i < tailRows.Length; i++)
                {
                    TestAssert.True(Math.Abs(tailRows[i - 1].Y - tailRows[i].Y - (i == 1 ? 1.15d : 1d) * size) < 0.02d,
                        "The outgoing mixed step uses both descents; later rows use the tail face alone.");
                }
                TestAssert.True(prefixRows.Concat(tailRows).All(show =>
                    show.X + show.Font.MeasureTextPoints(show.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d),
                    "Visible words must stay inside the balloon.");
                TestAssert.Equal(" ", shows[^1].Text);
                TestAssert.Equal(500, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
                TestAssert.True(Math.Abs(shows[^1].Size - size) < 0.001d, "The terminal mark keeps nominal balloon size.");
                TestAssert.True(Math.Abs(balloon.Height - (12.61d + prefixRows[0].Y - tailRows[^1].Y)) < 0.02d,
                    "Height includes prefix-only, incoming mixed, outgoing mixed and tail-only steps.");

                foreach ((string guardedPrefix, string guardedTail) in new[]
                {
                    (word + ";;", "ok. ending."), (prefix.Replace("table", "table,"), tail)
                })
                {
                    DocxParagraph guardedComment = comment with
                    {
                        Runs = [comment.Runs[0] with { Text = guardedPrefix }, comment.Runs[1] with { Text = guardedTail }]
                    };
                    DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(guardedComment)] }] };
                    var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) &&
                            !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Overwide prefix words and punctuation retain the complete fallback path.");
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonRetainsParagraphMarkFace()
    {
        foreach (string markFamily in new[] { "CommentFace", "OtherCommentFace" })
        foreach ((string prefix, string tail, bool admitted) in new[]
        {
            ("Review ", "table control.", true),
            ("Review ", string.Join(" ", Enumerable.Repeat("alpha", 30)), true),
            (string.Join(" ", Enumerable.Repeat("Review", 30)) + " ", "table control.", false)
        })
        {
            using MemoryStream stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                      <Default Extension="xml" ContentType="application/xml"/>
                      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                      <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                    </Types>
                    """,
                ["_rels/.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                    </Relationships>
                    """,
                ["word/_rels/document.xml.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                    </Relationships>
                    """,
                ["word/document.xml"] = """
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
                      <w:p><w:commentRangeStart w:id="1"/><w:r><w:rPr><w:rFonts w:ascii="AnchorFace" w:hAnsi="AnchorFace"/><w:sz w:val="24"/></w:rPr><w:t>Body anchor</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>
                      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="4140" w:bottom="1440" w:left="1440"/></w:sectPr>
                    </w:body></w:document>
                    """,
                ["word/comments.xml"] = $$"""
                    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer" w:initials="RV"><w:p>
                      <w:pPr><w:rPr><w:rFonts w:ascii="{{markFamily}}" w:hAnsi="{{markFamily}}"/><w:sz w:val="36"/></w:rPr></w:pPr>
                      <w:r><w:rPr><w:rFonts w:ascii="CommentFace" w:hAnsi="CommentFace"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{{prefix}}</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="OtherCommentFace" w:hAnsi="OtherCommentFace"/><w:sz w:val="24"/></w:rPr><w:t>{{tail}}</w:t></w:r>
                    </w:p></w:comment></w:comments>
                    """
            });
            DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null,
                CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
            var paragraph = ((DocxParagraphElement)document.RelatedStories.Single().BodyElements.Single()).Paragraph;
            DocxTextRun mark = paragraph.ParagraphMarkRun ?? throw new InvalidOperationException("The mixed comment paragraph mark must be retained.");
            TestAssert.Equal(" ", mark.Text);
            TestAssert.Equal(markFamily, mark.FontFamily);
            TestAssert.Equal(string.Empty, ((DocxParagraphElement)document.BodyElements[0]).Paragraph.ParagraphMarkRun!.Text);
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X).ToArray();
            var body = shows.Where(show => !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(prefix + tail, string.Join(" ", body.Select(show => show.Text.Trim())));
            var final = body[^1];
            var blank = shows.Single(show => show.Text == " " && Math.Abs(show.Y - final.Y) < 0.01d);
            TestAssert.Equal(admitted ? (markFamily == "CommentFace" ? 700 : 900) : 500,
                (int)blank.Font.Font.GetAdvanceWidth(blank.Font.Font.MapCodePoint(' ')));
            TestAssert.True(Math.Abs(blank.Size - final.Size) < 0.001d, "Declared mark size must not change nominal balloon size.");
            TestAssert.True(Math.Abs(blank.X - final.X - final.Font.MeasureTextPoints(final.Text, final.Size)) < 0.04d,
                "The final blank must follow the last visible source-face advance.");
            foreach (var intermediate in shows.Where(show => show.Text == " " && show.Y > final.Y + 0.01d))
            {
                TestAssert.Equal(admitted ? 900 : 500, (int)intermediate.Font.Font.GetAdvanceWidth(intermediate.Font.Font.MapCodePoint(' ')));
            }
        }
    }
}
