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
    public static void DocxWordCompatibleMixedBalloonKeepsFittingFourthBodyFaces()
    {
        foreach (int mask in Enumerable.Range(0, 8))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace"];
            var paragraphs = Enumerable.Range(0, 3).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = supplementary ? "a\U0001F600b." : "Review.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 3; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            TestAssert.Equal(3, last.Length);
            TestAssert.Equal(prefix, last[0].Text);
            TestAssert.Equal(suffix, last[1].Text);
            TestAssert.Equal(" ", last[2].Text);
            TestAssert.Equal(reverseFaces ? 900 : 700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[2].Font.Font.GetAdvanceWidth(last[2].Font.Font.MapCodePoint(' ')));
            for (int index = 0; index < 2; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Fourth-row body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[2].Size - last[0].Size) < 0.001d, "The fourth mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(2).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The fourth row uses the preceding body's descent and the larger body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and all three adjacent transitions.");
            TestAssert.True(last[2].X <= balloon.X + balloon.Width, "The fitting mixed fourth body stays within visible width.");
            foreach (DocxParagraph fourth in new[]
            {
                paragraphs[3] with { Runs = [.. paragraphs[3].Runs, paragraphs[3].Runs[0] with { Text = " extra." }] },
                paragraphs[3] with { Runs = [paragraphs[3].Runs[0] with { Bold = true }, paragraphs[3].Runs[1]] },
                paragraphs[3] with { Runs = [paragraphs[3].Runs[0] with { Text = new string('m', 120) }, paragraphs[3].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[3] with { Spacing = paragraphs[3].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(3), new DocxParagraphElement(fourth)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three fourth-body runs, decoration, an overwide closing run and unqualified spacing retain fallback.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 3
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or fourth-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingFifthBodyFaces()
    {
        foreach (int mask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace", "OtherCommentFace"];
            var paragraphs = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = supplementary ? "a\U0001F600b." : "Review.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 4; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            TestAssert.Equal(3, last.Length);
            TestAssert.Equal(prefix, last[0].Text);
            TestAssert.Equal(suffix, last[1].Text);
            TestAssert.Equal(" ", last[2].Text);
            TestAssert.Equal(reverseFaces ? 900 : 700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[2].Font.Font.GetAdvanceWidth(last[2].Font.Font.MapCodePoint(' ')));
            for (int index = 0; index < 2; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Fifth-row body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[2].Size - last[0].Size) < 0.001d, "The fifth mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(2).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The fifth row uses the preceding body's descent and the larger body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and all four adjacent transitions.");
            TestAssert.True(last[2].X <= balloon.X + balloon.Width, "The fitting mixed fifth body stays within visible width.");
            foreach (DocxParagraph fifth in new[]
            {
                paragraphs[4] with { Runs = [.. paragraphs[4].Runs, paragraphs[4].Runs[0] with { Text = " extra." }] },
                paragraphs[4] with { Runs = [paragraphs[4].Runs[0] with { Bold = true }, paragraphs[4].Runs[1]] },
                paragraphs[4] with { Runs = [paragraphs[4].Runs[0] with { Text = new string('m', 120) }, paragraphs[4].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[4] with { Spacing = paragraphs[4].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(fifth)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three fifth-body runs, decoration, an overwide closing run and unqualified spacing retain fallback.");
                if (fifth.Runs[0].Text.Length == 120 && mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
                {
                    var retainedStory = guardedStory with
                    {
                        BodyElements = guardedStory.BodyElements.Cast<DocxParagraphElement>()
                            .Select(element => (DocxBodyElement)(element with
                            {
                                Paragraph = element.Paragraph with
                                {
                                    ParagraphMarkRun = element.Paragraph.ParagraphMarkRun! with { FontFamily = "UnusedMarkFace" }
                                }
                            })).ToArray()
                    };
                    var formerStory = retainedStory with
                    {
                        BodyElements = retainedStory.BodyElements.Cast<DocxParagraphElement>()
                            .Select(element => (DocxBodyElement)(element with
                            {
                                Paragraph = element.Paragraph with { ParagraphMarkRun = null }
                            })).ToArray()
                    };
                    TestAssert.True(PaintedBalloonText(renderer, guarded with { RelatedStories = [retainedStory] }).SequenceEqual(
                        PaintedBalloonText(renderer, guarded with { RelatedStories = [formerStory] })),
                        "Unused marks preserve complete painted fallback text, advances and placement.");
                }
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 4
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or fifth-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingThirdBodyFaces()
    {
        foreach (int mask in Enumerable.Range(0, 4))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace"];
            var paragraphs = Enumerable.Range(0, 2).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = supplementary ? "a\U0001F600b." : "Review.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 2; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            TestAssert.Equal(3, last.Length);
            TestAssert.Equal(prefix, last[0].Text);
            TestAssert.Equal(suffix, last[1].Text);
            TestAssert.Equal(" ", last[2].Text);
            TestAssert.Equal(reverseFaces ? 900 : 700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[2].Font.Font.GetAdvanceWidth(last[2].Font.Font.MapCodePoint(' ')));
            for (int index = 0; index < 2; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Third-row body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[2].Size - last[0].Size) < 0.001d, "The third mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(2).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The third row uses the preceding body's descent and the larger body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and both adjacent transitions.");
            TestAssert.True(last[2].X <= balloon.X + balloon.Width, "The fitting mixed third body stays within visible width.");
            foreach (DocxParagraph fourth in new[]
            {
                paragraphs[2] with { Runs = [.. paragraphs[2].Runs, paragraphs[2].Runs[0] with { Text = " extra." }] },
                paragraphs[2] with { Runs = [paragraphs[2].Runs[0] with { Bold = true }, paragraphs[2].Runs[1]] },
                paragraphs[2] with { Runs = [paragraphs[2].Runs[0] with { Text = new string('m', 120) }, paragraphs[2].Runs[1] with { Text = " " + new string('m', 120) }] },
                paragraphs[2] with { Runs = [paragraphs[2].Runs[0] with { Text = new string('m', 120) }, paragraphs[2].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[2] with { Spacing = paragraphs[2].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(2), new DocxParagraphElement(fourth)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three third-body runs, decoration, an overwide closing run and unqualified spacing retain fallback.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 2
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or third-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsFittingSecondBodyFaces()
    {
        foreach (int mask in Enumerable.Range(0, 2))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace"];
            var paragraphs = Enumerable.Range(0, 1).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = supplementary ? "a\U0001F600b." : "Review.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 1; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                TestAssert.True(operations.Count >= 2, "The preceding paragraph retains separate body and nominal-mark emission.");
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            TestAssert.Equal(3, last.Length);
            TestAssert.Equal(prefix, last[0].Text);
            TestAssert.Equal(suffix, last[1].Text);
            TestAssert.Equal(" ", last[2].Text);
            TestAssert.Equal(reverseFaces ? 900 : 700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[2].Font.Font.GetAdvanceWidth(last[2].Font.Font.MapCodePoint(' ')));
            for (int index = 0; index < 2; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Second-row body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[2].Size - last[0].Size) < 0.001d, "The second mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(2).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The second row uses the preceding body's descent and the larger body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and the single adjacent transition.");
            TestAssert.True(last[2].X <= balloon.X + balloon.Width, "The fitting mixed second body stays within visible width.");
            foreach (DocxParagraph terminal in new[]
            {
                paragraphs[1] with { Runs = [.. paragraphs[1].Runs, paragraphs[1].Runs[0] with { Text = " extra." }] },
                paragraphs[1] with { Runs = [paragraphs[1].Runs[0] with { Bold = true }, paragraphs[1].Runs[1]] },
                paragraphs[1] with { Runs = [paragraphs[1].Runs[0] with { Text = new string('m', 120) }, paragraphs[1].Runs[1] with { Text = " " + new string('m', 120) }] },
                paragraphs[1] with { Runs = [paragraphs[1].Runs[0] with { Text = new string('m', 120) }, paragraphs[1].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[1] with { Spacing = paragraphs[1].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(1), new DocxParagraphElement(terminal)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three second-body runs, decoration, an overwide closing run and unqualified spacing retain fallback.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 1
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or second-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsSecondBodyBeforeFittingClosingRun()
    {
        foreach (int prefixLength in new[] { 80, 88, 96 })
        foreach (int mask in Enumerable.Range(0, 2))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace"];
            var paragraphs = Enumerable.Range(0, 1).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = (supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", prefixLength / 2)) : new string('a', prefixLength)) + " finish.";
            const string suffix = "tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 1; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                TestAssert.True(operations.Count >= 2, "The preceding paragraph retains separate body and nominal-mark emission.");
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            TestAssert.True(last.Length is 2 or 3, "The final row retains its body runs and nominal mark.");
            bool closingOwnRow = last.Length == 2;
            var secondShows = shows.Skip(operationIndex).ToArray();
            var secondRows = secondShows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(secondRows.Length >= 2, "The second first-body run must wrap before its fitting closing run.");
            TestAssert.Equal(prefix + (closingOwnRow ? " " : string.Empty), string.Concat(secondShows.Take(secondShows.Length - 2).Select(show => show.Text)));
            var firstSecond = secondRows[0].First();
            var firstFont = firstSecond.Font.Font;
            double ownPitch = (firstFont.Hhea.HorizontalAscender - firstFont.Hhea.HorizontalDescender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm * firstSecond.Size;
            for (int rowIndex = 1; rowIndex < secondRows.Length - 1; rowIndex++)
            {
                TestAssert.True(Math.Abs(secondRows[rowIndex - 1].Key - secondRows[rowIndex].Key - ownPitch) < 0.02d,
                    "Ordinary second continuations use only their first prepared body face.");
            }
            var precedingFont = previousParagraphFont!.Font;
            double firstStep = (-precedingFont.Hhea.HorizontalDescender / (double)precedingFont.UnitsPerEm +
                (firstFont.Hhea.HorizontalAscender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm) * firstSecond.Size;
            TestAssert.True(Math.Abs(previousParagraphLastY - firstSecond.Y - firstStep) < 0.02d,
                "The second first row uses its own first face, independently of the closing face.");
            if (closingOwnRow)
            {
                TestAssert.Equal(suffix.TrimStart(' '), last[0].Text);
                TestAssert.Equal(" ", last[1].Text);
                TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
                var preceding = secondRows[^2].ToArray();
                TestAssert.Equal(" ", preceding[^1].Text);
                TestAssert.Equal(last[0].Font, preceding[^1].Font);
                TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                    "A breakable preserved separator keeps the preceding last word on its original row.");
            }
            else
            {
                TestAssert.Equal(suffix, last[1].Text);
                TestAssert.Equal(" ", last[2].Text);
                TestAssert.Equal(reverseFaces ? 900 : 700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
                TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            }
            TestAssert.Equal(900, (int)last[^1].Font.Font.GetAdvanceWidth(last[^1].Font.Font.MapCodePoint(' ')));
            for (int index = 0; index < last.Length - 1; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Terminal body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[^1].Size - last[0].Size) < 0.001d, "The second mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(last.Length - 1).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The final mixed row uses its preceding first-body descent and the larger body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and the single adjacent transition.");
            TestAssert.True(last[^1].X <= balloon.X + balloon.Width, "The final mixed second row stays within visible width.");
            foreach (DocxParagraph second in new[]
            {
                paragraphs[1] with { Runs = [.. paragraphs[1].Runs, paragraphs[1].Runs[0] with { Text = " extra." }] },
                paragraphs[1] with { Runs = [paragraphs[1].Runs[0] with { Bold = true }, paragraphs[1].Runs[1]] },
                paragraphs[1] with { Runs = [paragraphs[1].Runs[0], paragraphs[1].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[1] with { Runs = [paragraphs[1].Runs[0] with { Text = prefix + " " }, paragraphs[1].Runs[1] with { Text = suffix.TrimStart(' ') }] },
                paragraphs[1] with { Spacing = paragraphs[1].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(1), new DocxParagraphElement(second)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three second-body runs, decoration, a closing run that cannot fit, a trailing prefix separator and unqualified spacing retain fallback.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 1
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or second-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsThirdBodyBeforeFittingClosingRun()
    {
        foreach (int prefixLength in new[] { 80, 88, 96 })
        foreach (int mask in Enumerable.Range(0, 4))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace"];
            var paragraphs = Enumerable.Range(0, 2).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = (supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", prefixLength / 2)) : new string('a', prefixLength)) + " finish.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 2; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            TestAssert.True(last.Length is 2 or 3, "The final row retains its body runs and nominal mark.");
            bool closingOwnRow = last.Length == 2;
            var thirdShows = shows.Skip(operationIndex).ToArray();
            var thirdRows = thirdShows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(thirdRows.Length >= 2, "The third first-body run must wrap before its fitting closing run.");
            TestAssert.Equal(prefix + (closingOwnRow ? " " : string.Empty), string.Concat(thirdShows.Take(thirdShows.Length - 2).Select(show => show.Text)));
            var firstThird = thirdRows[0].First();
            var firstFont = firstThird.Font.Font;
            double ownPitch = (firstFont.Hhea.HorizontalAscender - firstFont.Hhea.HorizontalDescender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm * firstThird.Size;
            for (int rowIndex = 1; rowIndex < thirdRows.Length - 1; rowIndex++)
            {
                TestAssert.True(Math.Abs(thirdRows[rowIndex - 1].Key - thirdRows[rowIndex].Key - ownPitch) < 0.02d,
                    "Ordinary third continuations use only their first prepared body face.");
            }
            var precedingFont = previousParagraphFont!.Font;
            double firstStep = (-precedingFont.Hhea.HorizontalDescender / (double)precedingFont.UnitsPerEm +
                (firstFont.Hhea.HorizontalAscender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm) * firstThird.Size;
            TestAssert.True(Math.Abs(previousParagraphLastY - firstThird.Y - firstStep) < 0.02d,
                "The third first row uses its own first face, independently of the closing face.");
            if (closingOwnRow)
            {
                TestAssert.Equal(suffix.TrimStart(' '), last[0].Text);
                TestAssert.Equal(" ", last[1].Text);
                TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
                var preceding = thirdRows[^2].ToArray();
                TestAssert.Equal(" ", preceding[^1].Text);
                TestAssert.Equal(last[0].Font, preceding[^1].Font);
                TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                    "A breakable preserved separator keeps the preceding last word on its original row.");
            }
            else
            {
                TestAssert.Equal(suffix, last[1].Text);
                TestAssert.Equal(" ", last[2].Text);
                TestAssert.Equal(reverseFaces ? 900 : 700, (int)last[0].Font.Font.GetAdvanceWidth(last[0].Font.Font.MapCodePoint('R')));
                TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[1].Font.Font.GetAdvanceWidth(last[1].Font.Font.MapCodePoint('R')));
            }
            TestAssert.Equal(900, (int)last[^1].Font.Font.GetAdvanceWidth(last[^1].Font.Font.MapCodePoint(' ')));
            for (int index = 0; index < last.Length - 1; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Terminal body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[^1].Size - last[0].Size) < 0.001d, "The third mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(last.Length - 1).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The final mixed row uses its preceding first-body descent and the larger body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and both adjacent transitions.");
            TestAssert.True(last[^1].X <= balloon.X + balloon.Width, "The final mixed third row stays within visible width.");
            foreach (DocxParagraph third in new[]
            {
                paragraphs[2] with { Runs = [.. paragraphs[2].Runs, paragraphs[2].Runs[0] with { Text = " extra." }] },
                paragraphs[2] with { Runs = [paragraphs[2].Runs[0] with { Bold = true }, paragraphs[2].Runs[1]] },
                paragraphs[2] with { Runs = [paragraphs[2].Runs[0], paragraphs[2].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[2] with { Runs = [paragraphs[2].Runs[0] with { Text = prefix + " " }, paragraphs[2].Runs[1] with { Text = suffix.TrimStart(' ') }] },
                paragraphs[2] with { Spacing = paragraphs[2].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(2), new DocxParagraphElement(third)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three third-body runs, decoration, a closing run that cannot fit, a trailing prefix separator and unqualified spacing retain fallback.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 2
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or third-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsFifthBodyBeforeFittingClosingRun()
    {
        foreach (int prefixLength in new[] { 80, 88, 104 })
        foreach (int mask in Enumerable.Range(0, 16))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace", "OtherCommentFace"];
            var paragraphs = Enumerable.Range(0, 4).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = (supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", prefixLength / 2)) : new string('a', prefixLength)) + " end.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 4; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            bool closingOwnRow = last.Length == 2;
            TestAssert.True(!closingOwnRow, "This fifth-terminal matrix joins a fitting closing run to the final first-body word.");
            TestAssert.True(last.Length is 2 or 3, "The final row retains a closing body and nominal mark.");
            var fifthShows = shows.Skip(operationIndex).ToArray();
            var fifthRows = fifthShows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(fifthRows.Length >= 2, "The fifth first-body run must wrap before its fitting closing run.");
            var firstFifth = fifthRows[0].First();
            var firstFont = firstFifth.Font.Font;
            TestAssert.Equal(prefix, string.Concat(fifthShows.Take(fifthShows.Length - 2)
                .Where(show => show.Font == firstFifth.Font).Select(show => show.Text)));
            double ownPitch = (firstFont.Hhea.HorizontalAscender - firstFont.Hhea.HorizontalDescender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm * firstFifth.Size;
            for (int rowIndex = 1; rowIndex < fifthRows.Length - 1; rowIndex++)
            {
                TestAssert.True(Math.Abs(fifthRows[rowIndex - 1].Key - fifthRows[rowIndex].Key - ownPitch) < 0.02d,
                    "Ordinary fifth continuations use only their first prepared body face.");
            }
            var precedingFont = previousParagraphFont!.Font;
            double firstStep = (-precedingFont.Hhea.HorizontalDescender / (double)precedingFont.UnitsPerEm +
                (firstFont.Hhea.HorizontalAscender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm) * firstFifth.Size;
            TestAssert.True(Math.Abs(previousParagraphLastY - firstFifth.Y - firstStep) < 0.02d,
                "The fifth first row uses its own first face, independently of the closing face.");
            TestAssert.Equal(closingOwnRow ? suffix.TrimStart(' ') : suffix, last[^2].Text);
            TestAssert.Equal(" ", last[^1].Text);
            TestAssert.Equal(reverseFaces ? 900 : 700, (int)firstFifth.Font.Font.GetAdvanceWidth(firstFifth.Font.Font.MapCodePoint('R')));
            TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[^2].Font.Font.GetAdvanceWidth(last[^2].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[^1].Font.Font.GetAdvanceWidth(last[^1].Font.Font.MapCodePoint(' ')));
            if (closingOwnRow)
            {
                var preceding = fifthRows[^2].ToArray();
                TestAssert.Equal(" ", preceding[^1].Text);
                TestAssert.Equal(last[0].Font, preceding[^1].Font);
                TestAssert.True(preceding[0].Text.EndsWith(" end.", StringComparison.Ordinal),
                    "The preserved separator remains after the preceding whole word in the closing face.");
                TestAssert.True(Math.Abs(last[0].X - firstFifth.X) < 0.001d,
                    "The closing-only fifth row resets to the continuation inset.");
            }
            for (int index = 0; index < last.Length - 1; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Fifth-row body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[^1].Size - last[0].Size) < 0.001d, "The fifth mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(last.Length - 1).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The final row uses preceding first-body descent and its emitted body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and all four adjacent transitions.");
            TestAssert.True(last[^1].X <= balloon.X + balloon.Width, "The final mixed fifth row stays within visible width.");
            foreach (DocxParagraph fifth in new[]
            {
                paragraphs[4] with { Runs = [.. paragraphs[4].Runs, paragraphs[4].Runs[0] with { Text = " extra." }] },
                paragraphs[4] with { Runs = [paragraphs[4].Runs[0] with { Bold = true }, paragraphs[4].Runs[1]] },
                paragraphs[4] with { Runs = [paragraphs[4].Runs[0], paragraphs[4].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[4] with { Runs = [paragraphs[4].Runs[0] with { Text = prefix + " " }, paragraphs[4].Runs[1] with { Text = suffix.TrimStart(' ') }] },
                paragraphs[4] with { Spacing = paragraphs[4].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(fifth)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three fifth-body runs, decoration, a closing run that cannot fit, a trailing prefix separator and unqualified spacing retain fallback.");
            }
            if (prefixLength == 80 && mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                // This former overflow fallback is now covered by Office's
                // closing-only row: keep the last first-body word on its row.
                var overflowFifth = paragraphs[4] with
                {
                    Runs = [paragraphs[4].Runs[0] with { Text = new string('a', 96) + " end." }, paragraphs[4].Runs[1]]
                };
                var overflowStory = story with { BodyElements = [.. story.BodyElements.Take(4), new DocxParagraphElement(overflowFifth)] };
                var overflowDocument = document with { RelatedStories = [overflowStory] };
                var overflowBalloon = renderer.InspectMarkupBalloons(overflowDocument).Single();
                var overflowShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(overflowDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= overflowBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                var overflowRows = overflowShows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                TestAssert.Equal(suffix.TrimStart(' '), overflowRows[^1].First().Text);
                TestAssert.True(overflowRows[^2].First().Text.EndsWith(" end.", StringComparison.Ordinal),
                    "The newly qualified closing-only row preserves the last first-body word instead of reflowing it.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 4
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or fifth-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsFourthBodyBeforeFittingClosingRun()
    {
        foreach (int prefixLength in new[] { 80, 88, 96 })
        foreach (int mask in Enumerable.Range(0, 8))
        foreach (bool supplementary in new[] { false, true })
        foreach (bool reverseFaces in new[] { false, true })
        foreach (short gap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        foreach (bool canonicalSpacing in new[] { false, true })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "CommentFace"];
            var paragraphs = Enumerable.Range(0, 3).Select(index =>
            {
                string text = (mask & (1 << index)) == 0 ? "end." : new string('m', 72) + " ending.";
                return DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, "AnchorFace")
                };
            }).ToList();
            string prefix = (supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", prefixLength / 2)) : new string('a', prefixLength)) + " finish.";
            const string suffix = " tail.";
            paragraphs.Add(paragraphs[0] with
            {
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, reverseFaces ? "OtherCommentFace" : "CommentFace"),
                    new DocxTextRun(suffix, 12d, null, false, false, false, null, reverseFaces ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = paragraphs[0].ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 30d }
            });
            if (canonicalSpacing)
            {
                paragraphs = paragraphs.Select(paragraph => paragraph with
                {
                    Spacing = paragraph.Spacing with
                    {
                        BeforeLinesValue = "300", AfterLinesValue = "300", BeforeAutoSpacingValue = "off", AfterAutoSpacingValue = "on",
                        BeforeValue = "120", AfterValue = "240"
                    }
                }).ToList();
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: gap, bodySpaceAdvance: 123),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && show.Text.Length != 0 && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            string expected = string.Join(" ", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))) + " ";
            TestAssert.Equal(expected, string.Concat(shows.Select(show => show.Text)));
            int operationIndex = 0;
            double previousParagraphLastY = shows[0].Y;
            PdfEmbeddedFont? previousParagraphFont = null;
            for (int paragraphIndex = 0; paragraphIndex < 3; paragraphIndex++)
            {
                string paragraphText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                var operations = new List<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)>();
                string joined = "";
                while (operationIndex < shows.Length && joined.Length < paragraphText.Length)
                {
                    operations.Add(shows[operationIndex]);
                    joined += shows[operationIndex++].Text;
                }
                TestAssert.Equal(paragraphText, joined);
                var bodyRows = operations.Take(operations.Count - 1).GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
                var embedded = bodyRows[0].First().Font;
                var font = embedded.Font;
                double pitch = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm * bodyRows[0].First().Size;
                for (int rowIndex = 1; rowIndex < bodyRows.Length; rowIndex++)
                {
                    TestAssert.True(Math.Abs(bodyRows[rowIndex - 1].Key - bodyRows[rowIndex].Key - pitch) < 0.02d,
                        "Each preceding paragraph retains its own continuation pitch.");
                }
                if (previousParagraphFont is not null)
                {
                    var previousBody = previousParagraphFont.Font;
                    double transition = (-previousBody.Hhea.HorizontalDescender / (double)previousBody.UnitsPerEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * bodyRows[0].First().Size;
                    TestAssert.True(Math.Abs(previousParagraphLastY - bodyRows[0].Key - transition) < 0.02d,
                        "Preceding paragraph boundaries retain their own prepared-face transition.");
                }
                previousParagraphLastY = bodyRows[^1].Key;
                previousParagraphFont = embedded;
            }
            var last = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Last().ToArray();
            bool closingOwnRow = last.Length == 2;
            TestAssert.True(last.Length is 2 or 3, "The final row retains a closing body and nominal mark.");
            var fourthShows = shows.Skip(operationIndex).ToArray();
            var fourthRows = fourthShows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).ToArray();
            TestAssert.True(fourthRows.Length >= 2, "The fourth first-body run must wrap before its fitting closing run.");
            var firstFourth = fourthRows[0].First();
            var firstFont = firstFourth.Font.Font;
            TestAssert.Equal(prefix, string.Concat(fourthShows.Take(fourthShows.Length - 2)
                .Where(show => show.Font == firstFourth.Font).Select(show => show.Text)));
            double ownPitch = (firstFont.Hhea.HorizontalAscender - firstFont.Hhea.HorizontalDescender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm * firstFourth.Size;
            for (int rowIndex = 1; rowIndex < fourthRows.Length - 1; rowIndex++)
            {
                TestAssert.True(Math.Abs(fourthRows[rowIndex - 1].Key - fourthRows[rowIndex].Key - ownPitch) < 0.02d,
                    "Ordinary fourth continuations use only their first prepared body face.");
            }
            var precedingFont = previousParagraphFont!.Font;
            double firstStep = (-precedingFont.Hhea.HorizontalDescender / (double)precedingFont.UnitsPerEm +
                (firstFont.Hhea.HorizontalAscender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm) * firstFourth.Size;
            TestAssert.True(Math.Abs(previousParagraphLastY - firstFourth.Y - firstStep) < 0.02d,
                "The fourth first row uses its own first face, independently of the closing face.");
            TestAssert.Equal(closingOwnRow ? suffix.TrimStart(' ') : suffix, last[^2].Text);
            TestAssert.Equal(" ", last[^1].Text);
            TestAssert.Equal(reverseFaces ? 900 : 700, (int)firstFourth.Font.Font.GetAdvanceWidth(firstFourth.Font.Font.MapCodePoint('R')));
            TestAssert.Equal(reverseFaces ? 700 : 900, (int)last[^2].Font.Font.GetAdvanceWidth(last[^2].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(900, (int)last[^1].Font.Font.GetAdvanceWidth(last[^1].Font.Font.MapCodePoint(' ')));
            if (closingOwnRow)
            {
                var preceding = fourthRows[^2].ToArray();
                TestAssert.Equal(" ", preceding[^1].Text);
                TestAssert.Equal(last[0].Font, preceding[^1].Font);
                TestAssert.True(preceding[0].Text.EndsWith(" finish.", StringComparison.Ordinal),
                    "The preserved separator remains after the preceding whole word in the closing face.");
                TestAssert.True(Math.Abs(last[0].X - firstFourth.X) < 0.001d,
                    "The closing-only fourth row resets to the continuation inset.");
            }
            for (int index = 0; index < last.Length - 1; index++)
            {
                TestAssert.True(Math.Abs(last[index + 1].X - last[index].X - last[index].Font.MeasureTextPoints(last[index].Text, last[index].Size)) < 0.04d,
                    "Fourth-row body parts and nominal mark follow their own advances.");
            }
            TestAssert.True(Math.Abs(last[^1].Size - last[0].Size) < 0.001d, "The fourth mark uses nominal balloon size.");
            var previous = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Reverse().Skip(1).First().First();
            var previousFont = previous.Font.Font;
            double nextAscent = last.Take(last.Length - 1).Max(show => (show.Font.Font.Hhea.HorizontalAscender + show.Font.Font.Hhea.HorizontalLineGap) / (double)show.Font.Font.UnitsPerEm);
            double step = (-previousFont.Hhea.HorizontalDescender / (double)previousFont.UnitsPerEm + nextAscent) * last[0].Size;
            TestAssert.True(Math.Abs(previous.Y - last[0].Y - step) < 0.02d,
                "The final row uses preceding first-body descent and its emitted body ascent/gap in either face order.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + shows[0].Y - last[0].Y)) < 0.03d,
                "Height and emission share preceding continuation pitches and all three adjacent transitions.");
            TestAssert.True(last[^1].X <= balloon.X + balloon.Width, "The final mixed fourth row stays within visible width.");
            foreach (DocxParagraph fourth in new[]
            {
                paragraphs[3] with { Runs = [.. paragraphs[3].Runs, paragraphs[3].Runs[0] with { Text = " extra." }] },
                paragraphs[3] with { Runs = [paragraphs[3].Runs[0] with { Bold = true }, paragraphs[3].Runs[1]] },
                paragraphs[3] with { Runs = [paragraphs[3].Runs[0], paragraphs[3].Runs[1] with { Text = new string('m', 120) }] },
                paragraphs[3] with { Runs = [paragraphs[3].Runs[0] with { Text = prefix + " " }, paragraphs[3].Runs[1] with { Text = suffix.TrimStart(' ') }] },
                paragraphs[3] with { Spacing = paragraphs[3].Spacing with { AfterLinesValue = "400" } }
            })
            {
                var guardedStory = story with { BodyElements = [.. story.BodyElements.Take(3), new DocxParagraphElement(fourth)] };
                var guarded = document with { RelatedStories = [guardedStory] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "Three fourth-body runs, decoration, a closing run that cannot fit, a trailing prefix separator and unqualified spacing retain fallback.");
            }
            if (mask == 0 && !supplementary && !reverseFaces && gap == 0 && rightMargin == 72d && !canonicalSpacing)
            {
                var numericParagraphs = paragraphs.Select((paragraph, index) => index < 3
                    ? paragraph with { Runs = [paragraph.Runs[0] with { FontFamily = "CommentFace", Text = "a." }] }
                    : paragraph).ToArray();
                var numericDocument = document with
                {
                    RelatedStories = [story with { BodyElements = numericParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                };
                int numericGuardIndex = 0;
                foreach (IFontResolver resolver in new IFontResolver[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var numericRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var numericBalloon = numericRenderer.InspectMarkupBalloons(numericDocument).Single();
                    var numericShows = ReadEmbeddedGlyphTextShows(numericRenderer.RenderPagesAsync(numericDocument, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= numericBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(numericBalloon.Height) && numericShows.Length > 0 &&
                        numericShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width, overwide-scalar and invalid leading or fourth-body pitches retain finite complete fallback: guard " + numericGuardIndex);
                    numericGuardIndex++;
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsParagraphsIndependently()
    {
        foreach (int paragraphCount in new[] { 2, 3, 4, 5, 6, 7 })
        foreach (int wrappedMask in paragraphCount == 5 ? Enumerable.Range(0, 32) : new[] { 0, 1 << (paragraphCount / 2), (1 << paragraphCount) - 1 })
        foreach (bool supplementary in new[] { false, true })
        foreach (short bodyGap in new short[] { 0, 450 })
        foreach (double rightMargin in new[] { 72d, 207d })
        {
            string[] families = ["CommentFace", "OtherCommentFace", "AnchorFace", "CommentFace", "OtherCommentFace"];
            string[] markFamilies = ["AnchorFace", "CommentFace", "OtherCommentFace", "AnchorFace", "CommentFace"];
            var paragraphs = new DocxParagraph[paragraphCount];
            for (int index = 0; index < paragraphs.Length; index++)
            {
                string word = supplementary ? string.Concat(Enumerable.Repeat("f\U0001F600", 40)) : new string('f', 80);
                string text = (wrappedMask & (1 << index)) != 0 ? word + " ending." : "end.";
                paragraphs[index] = DocxTests.CreateDocxLayoutParagraph(text, 12d, 14d) with
                {
                    LineSpacingPoints = null,
                    Runs = [new DocxTextRun(text, 12d, null, false, false, false, null, families[index % families.Length])],
                    ParagraphMarkRun = new DocxTextRun(" ", 18d + index * 6d, null, false, false, false, null, markFamilies[index % markFamilies.Length])
                };
            }
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: bodyGap, bodySpaceAdvance: 123,
                bodyDescender: -50, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.Text.Length != 0 && show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(string.Concat(paragraphs.Select(paragraph => paragraph.Runs[0].Text + " ")),
                string.Concat(shows.Select(show => show.Text)));
            int offset = 0;
            double expectedHeight = 12.61d;
            double previousY = 0d, previousDescentEm = 0d;
            for (int paragraphIndex = 0; paragraphIndex < paragraphCount; paragraphIndex++)
            {
                int start = offset;
                string joined = string.Empty;
                string expectedText = paragraphs[paragraphIndex].Runs[0].Text + " ";
                while (offset < shows.Length && joined != expectedText) { joined += shows[offset++].Text; }
                TestAssert.Equal(expectedText, joined);
                var paragraphShows = shows.Skip(start).Take(offset - start).ToArray();
                var mark = paragraphShows[^1];
                var bodyRows = paragraphShows.Take(paragraphShows.Length - 1)
                    .GroupBy(show => show.Y).Select(group => new
                    {
                        Text = string.Concat(group.Select(show => show.Text)),
                        group.First().X, group.First().Y, group.First().Size, group.First().Font
                    }).ToArray();
                TestAssert.True(bodyRows.Length >= ((wrappedMask & (1 << paragraphIndex)) != 0 ? 2 : 1),
                    "Each paragraph body fits or wraps independently.");
                var firstRow = bodyRows[0];
                var lastRow = bodyRows[^1];
                int expectedAdvance = families[paragraphIndex % families.Length] == "CommentFace" ? 700 : families[paragraphIndex % families.Length] == "OtherCommentFace" ? 900 : 500;
                TestAssert.Equal(expectedAdvance, (int)firstRow.Font.Font.GetAdvanceWidth(firstRow.Font.Font.MapCodePoint('R')));
                double size = firstRow.Size;
                var font = firstRow.Font.Font;
                double pitch = (-font.Hhea.HorizontalDescender + font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) * size / font.UnitsPerEm;
                for (int row = 1; row < bodyRows.Length; row++)
                {
                    TestAssert.True(Math.Abs(bodyRows[row - 1].Y - bodyRows[row].Y - pitch) < 0.02d,
                        "Every continuation uses its own prepared body's pitch.");
                }
                if (paragraphIndex > 0)
                {
                    double transition = (previousDescentEm +
                        (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm) * size;
                    TestAssert.True(Math.Abs(previousY - firstRow.Y - transition) < 0.02d,
                        "All adjacent transitions use previous descent and next ascent/gap.");
                    expectedHeight += transition;
                }
                expectedHeight += (bodyRows.Length - 1) * pitch;
                previousY = lastRow.Y;
                previousDescentEm = -font.Hhea.HorizontalDescender / (double)font.UnitsPerEm;
                int expectedMarkAdvance = markFamilies[paragraphIndex % markFamilies.Length] == "CommentFace" ? 123 : markFamilies[paragraphIndex % markFamilies.Length] == "OtherCommentFace" ? 900 : 500;
                TestAssert.Equal(" ", mark.Text);
                TestAssert.Equal(expectedMarkAdvance, (int)mark.Font.Font.GetAdvanceWidth(mark.Font.Font.MapCodePoint(' ')));
                TestAssert.True(Math.Abs(mark.Size - size) < 0.001d && Math.Abs(mark.Y - lastRow.Y) < 0.001d &&
                    Math.Abs(mark.X - lastRow.X - lastRow.Font.MeasureTextPoints(lastRow.Text, size)) < 0.04d,
                    "Each nominal mark follows its own final body row and prepared advance.");
                foreach (var row in bodyRows)
                {
                    TestAssert.True(!row.Text.Contains('\uFFFD') &&
                        row.X + row.Font.MeasureTextPoints(row.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d,
                        "Unicode progress and visible widths remain bounded in every paragraph.");
                }
            }
            TestAssert.Equal(shows.Length, offset);

            TestAssert.True(Math.Abs(balloon.Height - expectedHeight) < 0.02d,
                "Inspection and emission share every continuation and all adjacent paragraph transitions.");
            if (paragraphCount == 5 && wrappedMask == 0 && !supplementary && bodyGap == 0 && rightMargin == 72d)
            {
                foreach (var resolver in new[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyFirstAdvance: 32700),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700),
                    new BalloonTypefaceFontResolver(tailDescender: 32700)
                })
                {
                    var guardedParagraphs = paragraphs.ToArray();
                    guardedParagraphs[0] = guardedParagraphs[0] with { Runs = [guardedParagraphs[0].Runs[0] with { Text = "aaa" }] };
                    DocxDocument guarded = document with
                    {
                        RelatedStories = [story with { BodyElements = guardedParagraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() }]
                    };
                    var invalidRenderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var guardedBalloon = invalidRenderer.InspectMarkupBalloons(guarded).Single();
                    var guardedRows = ReadEmbeddedGlyphTextShows(invalidRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(guardedBalloon.Height) && guardedRows.Length > 0 &&
                        guardedRows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width and overwide bodies and nonpositive five-paragraph pitches retain finite complete fallback.");
                }

                foreach (DocxBodyElement[] guardedElements in new DocxBodyElement[][]
                {
                    paragraphs.Select((paragraph, index) => (DocxBodyElement)new DocxParagraphElement(index == 4 ? paragraph with { Runs = [paragraph.Runs[0] with { Bold = true }] } : paragraph)).ToArray(),
                    paragraphs.Select((paragraph, index) => (DocxBodyElement)new DocxParagraphElement(index == 4 ? paragraph with { Spacing = paragraph.Spacing with { AfterLinesValue = "400" } } : paragraph)).ToArray()
                })
                {
                    DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = guardedElements }] };
                    var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                    var guardedRows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedRows.Length > 0 && guardedRows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "decorated fifth bodies and unsupported spacing retain complete fallback.");
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonWrapsParagraphBodiesIndependently()
    {
        foreach (bool hasFourth in new[] { false, true })
        foreach (bool hasThird in new[] { false, true })
        foreach (bool wrappedFourth in new[] { false, true })
        foreach (bool wrappedThird in new[] { false, true })
        foreach (bool wrappedFirst in new[] { false, true })
        foreach (bool wrappedSecond in new[] { false, true })
        foreach (bool supplementary in new[] { false, true })
        foreach (bool zeroFirst in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (int length in new[] { 32, 80 })
        foreach (string firstMarkFace in new[] { "AnchorFace", "CommentFace" })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            if (wrappedFourth && (!hasFourth || !(wrappedFirst || wrappedSecond || wrappedThird))) { continue; }
            if (hasFourth && (!hasThird || !(wrappedFirst || wrappedSecond || wrappedThird))) { continue; }
            if (!hasThird && wrappedThird) { continue; }
            if (!wrappedFirst && !wrappedSecond && !wrappedThird) { continue; }
            AssertIndependentBalloonParagraphBodies(hasFourth, hasThird, wrappedFourth, wrappedThird, wrappedFirst, wrappedSecond, supplementary, zeroFirst, prefixDescender, prefixGap, length, firstMarkFace, rightMargin);
        }
    }

    public static void DocxWordCompatibleMixedBalloonParagraphBodyBoundaries()
    {
        for (int mask = 1; mask < 8; mask++)
        foreach (bool supplementary in new[] { false, true })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            bool fourth = (mask & 1) != 0;
            AssertIndependentBalloonParagraphBodies(fourth, fourth || (mask & 4) != 0, fourth && mask == 3,
                (mask & 4) != 0, (mask & 1) != 0, (mask & 2) != 0, supplementary, (mask & 2) != 0,
                (short)(mask % 2 == 0 ? -50 : -350), (short)(mask % 3 == 0 ? 0 : 450), mask % 2 == 0 ? 32 : 80,
                mask % 2 == 0 ? "AnchorFace" : "CommentFace", rightMargin);
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsLeadingFittingParagraphsAndWrapsFourth()
    {
        foreach (bool fourthParagraph in new[] { false, true })
        foreach (bool wrappedFourth in new[] { false, true })
        foreach (bool supplementary in new[] { false, true })
        foreach (bool middleSameFace in new[] { false, true })
        foreach (bool lastSameFace in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (string markFace in new[] { "AnchorFace", "OtherCommentFace" })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            if (!fourthParagraph && wrappedFourth) { continue; }
            string middleText = supplementary ? "a\U0001F600b" : "control.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("Review table", 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun("Review table", 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, markFace)
            };
            DocxParagraph middle = first with
            {
                Runs = [new DocxTextRun(middleText, 12d, null, false, false, false, null, middleSameFace ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "AnchorFace", FontSize = 24d }
            };
            DocxParagraph last = first with
            {
                Runs = [new DocxTextRun("ending.", 12d, null, false, false, false, null, lastSameFace ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "OtherCommentFace", FontSize = 36d }
            };
            string fourthText = wrappedFourth
                ? (supplementary ? string.Concat(Enumerable.Repeat("a\U0001F600", 40)) : new string('a', 80)) + " finish."
                : supplementary ? "a\U0001F600b" : "finish.";
            DocxParagraph fourth = first with
            {
                Runs = [first.Runs[0] with { Text = fourthText, FontFamily = "AnchorFace" }],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "CommentFace", FontSize = 48d }
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                fourthParagraph
                    ? [new DocxParagraphElement(first), new DocxParagraphElement(middle), new DocxParagraphElement(last), new DocxParagraphElement(fourth)]
                    : [new DocxParagraphElement(first), new DocxParagraphElement(middle), new DocxParagraphElement(last)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            if (!wrappedFourth) { TestAssert.Equal(fourthParagraph ? 8 : 6, shows.Length); }
            TestAssert.Equal("Review table " + middleText + " ending. " + (fourthParagraph ? fourthText + " " : string.Empty), string.Concat(shows.Select(show => show.Text)));
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(middleSameFace ? 700 : 900, (int)shows[2].Font.Font.GetAdvanceWidth(shows[2].Font.Font.MapCodePoint('b')));
            TestAssert.Equal(lastSameFace ? 700 : 900, (int)shows[4].Font.Font.GetAdvanceWidth(shows[4].Font.Font.MapCodePoint('b')));
            TestAssert.Equal(markFace == "AnchorFace" ? 500 : 900, (int)shows[1].Font.Font.GetAdvanceWidth(shows[1].Font.Font.MapCodePoint(' ')));
            TestAssert.Equal(500, (int)shows[3].Font.Font.GetAdvanceWidth(shows[3].Font.Font.MapCodePoint(' ')));
            TestAssert.Equal(900, (int)shows[5].Font.Font.GetAdvanceWidth(shows[5].Font.Font.MapCodePoint(' ')));
            double size = shows[0].Size;
            double firstPitch = -prefixDescender / 1000d + (middleSameFace ? 0.5d + prefixGap / 1000d : 0.8d);
            double secondPitch = (middleSameFace ? -prefixDescender / 1000d : 0.2d) + (lastSameFace ? 0.5d + prefixGap / 1000d : 0.8d);
            double thirdPitch = fourthParagraph ? (lastSameFace ? -prefixDescender / 1000d : 0.2d) + 0.8d : 0d;
            TestAssert.True(Math.Abs(shows[0].Y - shows[2].Y - firstPitch * size) < 0.02d &&
                Math.Abs(shows[2].Y - shows[4].Y - secondPitch * size) < 0.02d,
                "Every paragraph transition uses its adjacent prepared body metrics.");
            int fourthRowsCount = 0;
            if (fourthParagraph)
            {
                var fourthRows = shows.Skip(6).Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
                fourthRowsCount = fourthRows.Length;
                TestAssert.True(fourthRowsCount >= (wrappedFourth ? 2 : 1), "The fourth body fits or wraps independently of its three leading paragraphs.");
                TestAssert.Equal(500, (int)shows[6].Font.Font.GetAdvanceWidth(shows[6].Font.Font.MapCodePoint('b')));
                TestAssert.Equal(123, (int)shows[^1].Font.Font.GetAdvanceWidth(shows[^1].Font.Font.MapCodePoint(' ')));
                TestAssert.True(Math.Abs(shows[4].Y - shows[6].Y - thirdPitch * size) < 0.02d && Math.Abs(shows[6].X - shows[4].X) < 0.001d,
                    "The fourth paragraph uses its own body and the adjacent third-body descent.");
                foreach (var row in fourthRows)
                {
                    TestAssert.Equal(500, (int)row.Font.Font.GetAdvanceWidth(row.Font.Font.MapCodePoint('b')));
                    TestAssert.True(Math.Abs(row.X - shows[4].X) < 0.001d && row.X + row.Font.MeasureTextPoints(row.Text.TrimEnd(), size) <= balloon.X + balloon.Width + 0.02d,
                        "Every fourth-body row uses its own face and fits the continuation inset.");
                }
                for (int index = 1; index < fourthRows.Length; index++)
                {
                    TestAssert.True(Math.Abs(fourthRows[index - 1].Y - fourthRows[index].Y - size) < 0.02d,
                        "Fourth-body continuations use only their own prepared metrics.");
                }
                TestAssert.Equal(" ", shows[^1].Text);
                TestAssert.True(Math.Abs(shows[^1].Y - fourthRows[^1].Y) < 0.001d && Math.Abs(shows[^1].Size - size) < 0.001d &&
                    Math.Abs(shows[^1].X - fourthRows[^1].X - fourthRows[^1].Font.MeasureTextPoints(fourthRows[^1].Text, size)) < 0.04d,
                    "The fourth mark retains nominal size and follows its own final row.");
            }
            foreach (int index in new[] { 0, 2, 4 })
            {
                TestAssert.Equal(" ", shows[index + 1].Text);
                TestAssert.True(Math.Abs(shows[index].Y - shows[index + 1].Y) < 0.001d && Math.Abs(shows[index + 1].Size - size) < 0.001d,
                    "Each mark stays beside its own paragraph at nominal balloon size.");
                TestAssert.True(Math.Abs(shows[index + 1].X - shows[index].X - shows[index].Font.MeasureTextPoints(shows[index].Text, size)) < 0.04d,
                    "Each mark follows its paragraph's visible advance.");
            }
            TestAssert.True(Math.Abs(shows[2].X - shows[4].X) < 0.001d && shows[2].X < shows[0].X,
                "Later paragraphs retain the continuation inset.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + (firstPitch + secondPitch + thirdPitch + Math.Max(0, fourthRowsCount - 1)) * size)) < 0.02d,
                "Height and emission share every paragraph transition.");
            if (!supplementary && !middleSameFace && !lastSameFace && prefixDescender == -50 && prefixGap == 0 && markFace == "AnchorFace")
            {
                foreach (DocxBodyElement[] guardedElements in new DocxBodyElement[][]
                {
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle), new DocxParagraphElement(last), new DocxParagraphElement(fourth with { Runs = [fourth.Runs[0] with { Bold = true }] })],
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle), new DocxParagraphElement(last), new DocxParagraphElement(fourth with { ParagraphMarkRun = fourth.ParagraphMarkRun! with { Bold = true } })],
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle with { Runs = [middle.Runs[0] with { Text = new string('a', 200) }] }), new DocxParagraphElement(last with { Runs = [last.Runs[0] with { Text = new string('a', 200), Bold = true }] })],
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle with { Spacing = middle.Spacing with { AfterLinesValue = "400" } }), new DocxParagraphElement(last)],
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle), new DocxParagraphElement(last with { Runs = [last.Runs[0] with { Text = new string('a', 200) }, first.Runs[0] with { Text = " " + new string('m', 200) }] })],
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle), new DocxParagraphElement(last with { ParagraphMarkRun = last.ParagraphMarkRun! with { Bold = true } })],
                    [new DocxParagraphElement(first), new DocxParagraphElement(middle with { Runs = [middle.Runs[0] with { Text = "two  words" }] }), new DocxParagraphElement(last)]
                })
                {
                    DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = guardedElements }] };
                    var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "explicit line spacing, mixed/decorated faces and repeated spaces retain complete fallback.");
                }
                if (wrappedFourth)
                {
                    foreach (var invalidResolver in new[]
                    {
                        new BalloonTypefaceFontResolver(tailFirstAdvance: 32700),
                        new BalloonTypefaceFontResolver(tailFirstAdvance: 0),
                        new BalloonTypefaceFontResolver(tailDescender: 32700)
                    })
                    {
                        DocxDocument guarded = document with
                        {
                            RelatedStories = [story with { BodyElements = [new DocxParagraphElement(first),
                                new DocxParagraphElement(middle with { Runs = [middle.Runs[0] with { FontFamily = "CommentFace" }] }),
                                new DocxParagraphElement(last with { Runs = [last.Runs[0] with { FontFamily = "CommentFace" }] }),
                                new DocxParagraphElement(fourth with { Runs = [fourth.Runs[0] with { Text = new string('a', 80), FontFamily = "OtherCommentFace" }] })] }]
                        };
                        var invalidRenderer = new DocxRenderer(invalidResolver, OoxPdfDocxMarkupMode.AllMarkup,
                            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                        var guardedBalloon = invalidRenderer.InspectMarkupBalloons(guarded).Single();
                        var guardedShows = ReadEmbeddedGlyphTextShows(invalidRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                            .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                        TestAssert.True(double.IsFinite(guardedBalloon.Height) && guardedShows.Length > 0 &&
                            guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                            "An oversized fourth scalar, zero-width fourth body or nonpositive fourth pitch retains finite complete fallback.");
                    }
                }
                foreach (var invalidResolver in new[]
                {
                    new BalloonTypefaceFontResolver(bodyAdvance: 0, bodySpaceAdvance: 0),
                    new BalloonTypefaceFontResolver(bodyAscender: -32700)
                })
                {
                    DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = [new DocxParagraphElement(first), new DocxParagraphElement(first), new DocxParagraphElement(first)] }] };
                    var invalidRenderer = new DocxRenderer(invalidResolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                    var guardedBalloon = invalidRenderer.InspectMarkupBalloons(guarded).Single();
                    var guardedShows = ReadEmbeddedGlyphTextShows(invalidRenderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                        .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                    TestAssert.True(double.IsFinite(guardedBalloon.Height) && guardedShows.Length > 0 &&
                        guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                        "Zero-width bodies and nonpositive adjacent metrics retain finite complete fallback.");
                }
            }
        }
    }

    public static void DocxWordCompatibleMixedBalloonNormalizesExplicitParagraphSpacing()
    {
        foreach (string? beforeAuto in new string?[] { null, "1", "0", "true", "false", "on", "off" })
        foreach (string? afterAuto in new string?[] { null, "1", "0", "true", "false", "on", "off" })
        foreach (string? beforeLines in new string?[] { null, "100", "200", "300" })
        foreach (string? afterLines in new string?[] { null, "100", "200", "300" })
        foreach (string value in new[] { "120", "0", "480", "4294967295" })
        foreach (int spacingKind in new[] { 0, 1, 2 })
        foreach (int bodyKind in new[] { 0, 1, 2 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            AssertMixedBalloonSpacing(beforeAuto, afterAuto, beforeLines, afterLines, value, spacingKind, bodyKind, rightMargin);
        }
    }

    public static void DocxWordCompatibleMixedBalloonSpacingBoundaries()
    {
        string?[] automatic = [null, "1", "0", "true", "false", "on", "off"];
        string?[] lines = [null, "100", "200", "300"];
        string[] points = ["120", "0", "480", "4294967295"];
        int index = 0;
        foreach (int spacingKind in new[] { 0, 1, 2 })
        foreach (int bodyKind in new[] { 0, 1, 2 })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            AssertMixedBalloonSpacing(automatic[index % 7], automatic[(index + 3) % 7],
                lines[index % 4], lines[(index + 1) % 4], points[index % 4], spacingKind, bodyKind, rightMargin);
            index++;
        }
    }

    public static void DocxWordCompatibleMixedBalloonKeepsTwoFittingParagraphsAndTheirMarks()
    {
        foreach (bool supplementary in new[] { false, true })
        foreach (bool sameFace in new[] { false, true })
        foreach (short prefixDescender in new short[] { -50, -350 })
        foreach (short prefixGap in new short[] { 0, 450 })
        foreach (string markFace in new[] { "AnchorFace", "OtherCommentFace" })
        foreach (double rightMargin in new[] { 72d, 144d, 207d })
        {
            const string prefix = "Review table";
            string tail = supplementary ? "a\U0001F600b" : "control.";
            DocxParagraph anchor = DocxTests.CreateCommentMarkerParagraph("Body anchor", "1") with
            {
                Runs = [new DocxTextRun("Body anchor", 12d, null, false, false, false, null, "AnchorFace")]
            };
            DocxParagraph first = DocxTests.CreateDocxLayoutParagraph(prefix, 12d, 14d) with
            {
                LineSpacingPoints = null,
                Runs = [new DocxTextRun(prefix, 12d, null, false, false, false, null, "CommentFace")],
                ParagraphMarkRun = new DocxTextRun(" ", 18d, null, false, false, false, null, markFace)
            };
            DocxParagraph second = first with
            {
                Runs = [new DocxTextRun(tail, 12d, null, false, false, false, null, sameFace ? "CommentFace" : "OtherCommentFace")],
                ParagraphMarkRun = first.ParagraphMarkRun! with { FontFamily = "AnchorFace", FontSize = 24d }
            };
            DocxRelatedStory story = new(DocxRelatedStoryKind.Comment, "/word/comments.xml", "1",
                [new DocxParagraphElement(first), new DocxParagraphElement(second)], [], [], null)
            {
                CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", null, null, null, null, null)
            };
            DocxDocument document = new(612d, 792d, 72d, rightMargin, 72d, 72d, DocxPageSettings.Empty,
                [], [], [], [new DocxParagraphElement(anchor)], [], [])
            {
                RelatedStories = [story], MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
            var renderer = new DocxRenderer(new BalloonTypefaceFontResolver(bodyLineGap: prefixGap, bodySpaceAdvance: 123,
                bodyDescender: prefixDescender, bodyAscender: 500), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var balloon = renderer.InspectMarkupBalloons(document).Single();
            var shows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single())
                .Where(show => show.X >= balloon.X && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(4, shows.Length);
            TestAssert.Equal(prefix, shows[0].Text);
            TestAssert.Equal(" ", shows[1].Text);
            TestAssert.Equal(tail, shows[2].Text);
            TestAssert.Equal(" ", shows[3].Text);
            TestAssert.Equal(700, (int)shows[0].Font.Font.GetAdvanceWidth(shows[0].Font.Font.MapCodePoint('R')));
            TestAssert.Equal(sameFace ? 700 : 900, (int)shows[2].Font.Font.GetAdvanceWidth(shows[2].Font.Font.MapCodePoint('b')));
            TestAssert.Equal(markFace == "AnchorFace" ? 500 : 900, (int)shows[1].Font.Font.GetAdvanceWidth(shows[1].Font.Font.MapCodePoint(' ')));
            TestAssert.Equal(500, (int)shows[3].Font.Font.GetAdvanceWidth(shows[3].Font.Font.MapCodePoint(' ')));
            double size = shows[0].Size;
            double pitch = -prefixDescender / 1000d + (sameFace ? 0.5d + prefixGap / 1000d : 0.8d);
            TestAssert.True(Math.Abs(shows[0].Y - shows[2].Y - pitch * size) < 0.02d,
                "Paragraph spacing uses the visible first face's descent and next face's ascent and gap.");
            foreach (int index in new[] { 0, 2 })
            {
                TestAssert.True(Math.Abs(shows[index].Y - shows[index + 1].Y) < 0.001d, "Each paragraph mark stays on its own row.");
                TestAssert.True(Math.Abs(shows[index + 1].Size - size) < 0.001d, "Paragraph marks use nominal balloon size.");
                TestAssert.True(Math.Abs(shows[index + 1].X - shows[index].X - shows[index].Font.MeasureTextPoints(shows[index].Text, size)) < 0.04d,
                    "Each paragraph mark follows its own visible advance.");
            }
            TestAssert.True(shows[2].X < shows[0].X, "The second paragraph starts at the continuation inset.");
            TestAssert.True(Math.Abs(balloon.Height - (12.61d + pitch * size)) < 0.02d, "Geometry and emission share the paragraph step.");
            foreach (DocxBodyElement[] guardedElements in new DocxBodyElement[][]
            {
                [new DocxParagraphElement(first with { Spacing = first.Spacing with { AfterLinesValue = "400" } }), new DocxParagraphElement(second)],
                [new DocxParagraphElement(first with { Runs = [first.Runs[0] with { Bold = true }] }), new DocxParagraphElement(second)]
            })
            {
                DocxDocument guarded = document with { RelatedStories = [story with { BodyElements = guardedElements }] };
                var guardedBalloon = renderer.InspectMarkupBalloons(guarded).Single();
                var guardedShows = ReadEmbeddedGlyphTextShows(renderer.RenderPagesAsync(guarded, null, CancellationToken.None).ToBlockingEnumerable().Single())
                    .Where(show => show.X >= guardedBalloon.X && !string.IsNullOrWhiteSpace(show.Text) && !show.Text.StartsWith("Commented", StringComparison.Ordinal)).ToArray();
                TestAssert.True(guardedShows.Length > 0 && guardedShows.All(show => show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint('R')) == 500),
                    "unsupported spacing and decorated runs retain complete fallback.");
            }
        }
    }
}
