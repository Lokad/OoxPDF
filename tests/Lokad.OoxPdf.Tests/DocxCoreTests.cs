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

internal static class DocxCoreTests
{
    public static void DocxReaderParsesOnOffRunProperties()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:b w:val="on"/><w:i w:val="off"/></w:rPr><w:t>OnOff</w:t></w:r>
                      <w:r><w:rPr><w:b w:val="off"/><w:i w:val="on"/></w:rPr><w:t>OffOn</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph paragraph = document.Paragraphs[0];
        TestAssert.True(paragraph.Runs[0].Bold, "Expected w:b w:val=\"on\" to enable bold.");
        TestAssert.True(paragraph.Runs[0].Italic == false, "Expected w:i w:val=\"off\" to disable italic.");
        TestAssert.True(paragraph.Runs[1].Bold == false, "Expected w:b w:val=\"off\" to disable bold.");
        TestAssert.True(paragraph.Runs[1].Italic, "Expected w:i w:val=\"on\" to enable italic.");
    }

    public static void DocxReaderParsesRunCharacterSpacing()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:spacing w:val="40"/></w:rPr><w:t>Wide</w:t></w:r>
                      <w:r><w:rPr><w:spacing w:val="-20"/></w:rPr><w:t>Tight</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph paragraph = document.Paragraphs[0];
        TestAssert.Equal(2d, paragraph.Runs[0].CharacterSpacingPoints);
        TestAssert.Equal(-1d, paragraph.Runs[1].CharacterSpacingPoints);
    }

    public static void DocxRendererEmitsRunCharacterSpacingAsPositionedGlyphs()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:sz w:val="24"/><w:spacing w:val="40"/></w:rPr><w:t>AB</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("TJ", pdf);
        TestAssert.Contains("-166.667", pdf);
    }

    public static void OfficePdfTextProfileExposesWordNumberedListCharacterSpacing()
    {
        TestAssert.Equal(0.048d, OfficePdfTextEmissionProfile.ObservedWordNumberedListTextStateCharacterSpacing(12d));
        TestAssert.Equal(0.044d, OfficePdfTextEmissionProfile.ObservedWordNumberedListTextStateCharacterSpacing(11d));
    }

    public static void DocxTextEmissionPlannerOwnsListLabelTextStateTarget()
    {
        var decimalLabel = new DocxListLabel("1", "decimal", "%1.", "tab", "7", 0, DocxNumberingIndent.Empty, DocxTextRunStyle.Empty);
        var bulletLabel = new DocxListLabel("*", "bullet", "\uF0B7", "tab", "7", 0, DocxNumberingIndent.Empty, DocxTextRunStyle.Empty);
        var decimalRun = new DocxTextRun("1", 12d, null, false, false, false, null, null);
        var bulletRun = new DocxTextRun("*", 12d, null, false, false, false, null, null);
        DocxTextStateCharacterSpacingTarget decimalTarget = DocxTextEmissionPlanner.TextStateCharacterSpacingTargetForListLabel(decimalLabel, 12d);
        DocxTextStateCharacterSpacingTarget bulletTarget = DocxTextEmissionPlanner.TextStateCharacterSpacingTargetForListLabel(bulletLabel, 12d);
        DocxTextEmissionPlan decimalPlan = DocxTextEmissionPlanner.CreateForListLabel(decimalRun, decimalLabel);
        DocxTextEmissionPlan bulletPlan = DocxTextEmissionPlanner.CreateForListLabel(bulletRun, bulletLabel);

        TestAssert.Equal(0.048d, DocxTextEmissionPlanner.TextStateCharacterSpacingForListLabel(decimalLabel, 12d));
        TestAssert.Equal(0d, DocxTextEmissionPlanner.TextStateCharacterSpacingForListLabel(bulletLabel, 12d));
        TestAssert.Equal(0.048d, decimalTarget.CharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, decimalTarget.Source);
        TestAssert.Equal(0d, bulletTarget.CharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, bulletTarget.Source);
        TestAssert.Equal(0.048d, decimalPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, decimalPlan.PdfCharacterSpacingSource);
        TestAssert.True(!decimalPlan.CompensatePdfCharacterSpacing, "List-label PDF text state should be emitted, not folded into layout positioning.");
        TestAssert.Equal(0d, bulletPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, bulletPlan.PdfCharacterSpacingSource);
    }

    public static void DocxTextEmissionPlannerOwnsPdfTextStateAndPositioningSpacing()
    {
        var run = new DocxTextRun("Tracked", 11d, null, false, false, false, null, null, 0.25d);

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.Create(run, 11d, pdfCharacterSpacing: 0.05d, compensatePdfCharacterSpacing: true, source: DocxTextStateCharacterSpacingSource.None);

        TestAssert.Equal(11.04d, plan.PdfFontSize);
        TestAssert.Equal(0.05d, plan.PdfCharacterSpacing);
        TestAssert.Equal(0.20d, plan.PositioningCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.Explicit, plan.PdfCharacterSpacingSource);
        TestAssert.True(plan.CompensatePdfCharacterSpacing, "Planner should record that positioned glyph advances compensate PDF Tc.");
    }

    public static void DocxTextEmissionPlannerKeepsNumberedLabelTcOutOfPositioning()
    {
        var run = new DocxTextRun("1", 12d, null, false, false, false, null, null);
        double numberedTc = OfficePdfTextEmissionProfile.ObservedWordNumberedListTextStateCharacterSpacing(12d);

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.Create(run, 12d, numberedTc, compensatePdfCharacterSpacing: false, source: DocxTextStateCharacterSpacingSource.None);

        TestAssert.Equal(12d, plan.PdfFontSize);
        TestAssert.Equal(numberedTc, plan.PdfCharacterSpacing);
        TestAssert.Equal(0d, plan.PositioningCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.Explicit, plan.PdfCharacterSpacingSource);
        TestAssert.True(!plan.CompensatePdfCharacterSpacing, "Numbered labels use PDF Tc as emitted text state, not as a compensated layout offset.");
    }

    public static void DocxTextEmissionPlannerForcesTerminalLineSpacesToNeutralTc()
    {
        var run = new DocxTextRun("Body", 11d, null, false, false, false, null, null, 0.25d);

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateTerminalLineSpace(run, 11d);

        TestAssert.Equal(11.04d, plan.PdfFontSize);
        TestAssert.Equal(0d, plan.PdfCharacterSpacing);
        TestAssert.Equal(0.25d, plan.PositioningCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.TerminalLineSpace, plan.PdfCharacterSpacingSource);
        TestAssert.True(plan.CompensatePdfCharacterSpacing, "Terminal spaces should stay eligible for authored positioning while emitting neutral PDF Tc.");
    }

    public static void DocxTextEmissionPlannerOwnsTerminalSegmentPlanOverride()
    {
        var run = new DocxTextRun("Body", 11d, null, false, false, false, null, null, 0.25d);

        DocxTextEmissionPlan textPlan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            run,
            11d,
            pdfCharacterSpacing: 0.05d,
            DocxTextStateCharacterSpacingSource.AdvanceTarget,
            compensatePdfCharacterSpacing: false,
            isTerminalLineSpace: false);
        DocxTextEmissionPlan terminalPlan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            run,
            11d,
            pdfCharacterSpacing: 0.05d,
            DocxTextStateCharacterSpacingSource.AdvanceTarget,
            compensatePdfCharacterSpacing: false,
            isTerminalLineSpace: true);

        TestAssert.Equal(0.05d, textPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, textPlan.PdfCharacterSpacingSource);
        TestAssert.True(!textPlan.CompensatePdfCharacterSpacing, "Non-terminal segments should preserve the caller's compensation mode.");
        TestAssert.Equal(0d, terminalPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.TerminalLineSpace, terminalPlan.PdfCharacterSpacingSource);
        TestAssert.True(terminalPlan.CompensatePdfCharacterSpacing, "Terminal spaces should always be planned through the neutral terminal-space path.");
        TestAssert.Equal(0.25d, terminalPlan.PositioningCharacterSpacing);
    }

    public static void DocxTextEmissionPlannerDerivesTcFromAdvanceTarget()
    {
        var run = new DocxTextRun("Body", 11d, null, false, false, false, null, null, 0.12d);

        DocxTextStateAdvanceTarget target = DocxTextEmissionPlanner.CreateAdvanceTarget(
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d);
        double tc = DocxTextEmissionPlanner.TextStateCharacterSpacingForAdvanceTarget(
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d);
        DocxTextEmissionPlan emittedPlan = DocxTextEmissionPlanner.CreateForAdvanceTarget(
            run,
            11d,
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d,
            compensatePdfCharacterSpacing: false);
        DocxTextEmissionPlan compensatedPlan = DocxTextEmissionPlanner.CreateForAdvanceTarget(
            run,
            11d,
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d,
            compensatePdfCharacterSpacing: true);

        TestAssert.Equal(3, target.GlyphGapCount);
        TestAssert.Equal(24d, target.CurrentEmittedAdvance);
        TestAssert.Equal(24.18d, target.TargetEmittedAdvance);
        TestAssert.True(Math.Abs(target.CharacterSpacing - 0.06d) < 0.0001d, "Advance target should expose the derived uniform text-state spacing.");
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, target.Source);
        TestAssert.True(Math.Abs(tc - 0.06d) < 0.0001d, "Tc should be the emitted-advance delta distributed over glyph gaps.");
        TestAssert.True(Math.Abs(emittedPlan.PdfCharacterSpacing - 0.06d) < 0.0001d, "Uncompensated plans should carry the derived Tc.");
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, emittedPlan.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(emittedPlan.PositioningCharacterSpacing - 0.12d) < 0.0001d, "Uncompensated plans should leave positioning spacing unchanged.");
        TestAssert.True(Math.Abs(compensatedPlan.PdfCharacterSpacing - 0.06d) < 0.0001d, "Compensated plans should carry the same derived Tc.");
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, compensatedPlan.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(compensatedPlan.PositioningCharacterSpacing - 0.06d) < 0.0001d, "Compensated plans should subtract derived Tc from positioning spacing.");
        TestAssert.Equal(0d, DocxTextEmissionPlanner.TextStateCharacterSpacingForAdvanceTarget(0, 24d, 24.18d));
    }

    public static void DocxTextEmissionPlannerSplitsDashPunctuationIntoOperationParts()
    {
        var run = new DocxTextRun("Alpha-Beta", 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout("Alpha-Beta", run, 20d, 100d, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text);

        IReadOnlyList<DocxTextEmissionPart> parts = DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, 10d, new DocxTests.FontSizeWidthTextMeasurer());

        TestAssert.Equal(3, parts.Count);
        TestAssert.Equal("Alpha", parts[0].Text);
        TestAssert.Equal("-", parts[1].Text);
        TestAssert.Equal("Beta", parts[2].Text);
        TestAssert.Equal(20d, parts[0].X);
        TestAssert.Equal(70d, parts[1].X);
        TestAssert.Equal(80d, parts[2].X);
        TestAssert.Equal(40d, parts[2].Width);
    }

    public static void DocxTextEmissionPlannerKeepsWholeOperationWithoutMeasurer()
    {
        var run = new DocxTextRun("Alpha-Beta", 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout("Alpha-Beta", run, 20d, 50d, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text);

        IReadOnlyList<DocxTextEmissionPart> parts = DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, 10d, null);

        TestAssert.Equal(1, parts.Count);
        TestAssert.Equal("Alpha-Beta", parts[0].Text);
        TestAssert.Equal(20d, parts[0].X);
        TestAssert.Equal(50d, parts[0].Width);
    }

    public static void DocxTextEmissionPlannerSkipsEmptyOperationParts()
    {
        var run = new DocxTextRun(string.Empty, 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout(string.Empty, run, 20d, 0d, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text);

        IReadOnlyList<DocxTextEmissionPart> parts = DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, 10d, new DocxTests.FontSizeWidthTextMeasurer());

        TestAssert.Equal(0, parts.Count);
    }

    public static void DocxTextEmissionPlannerClassifiesTextWithoutExposingIt()
    {
        DocxTextEmissionCharacterProfile profile = DocxTextEmissionPlanner.ClassifyText("A9 -+");

        TestAssert.Equal(1, profile.LetterCount);
        TestAssert.Equal(1, profile.DigitCount);
        TestAssert.Equal(1, profile.WhitespaceCount);
        TestAssert.Equal(1, profile.PunctuationCount);
        TestAssert.Equal(1, profile.SymbolCount);
        TestAssert.Equal(0, profile.OtherCount);
    }

    public static void DocxReaderPreservesDocumentSettingsCompatibilityFacts()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
                </Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:defaultTabStop w:val="720"/>
                  <w:characterSpacingControl w:val="doNotCompress"/>
                  <w:revisionView w:markup="1" w:comments="0" w:insDel="1" w:formatting="0" w:inkAnnotations="1"/>
                  <w:trackRevisions/>
                  <w:doNotTrackMoves w:val="0"/>
                  <w:doNotTrackFormatting w:val="1"/>
                  <w:mirrorMargins/>
                  <w:footnotePr>
                    <w:pos w:val="pageBottom"/>
                    <w:numFmt w:val="lowerRoman"/>
                    <w:numStart w:val="4"/>
                    <w:numRestart w:val="continuous"/>
                  </w:footnotePr>
                  <w:endnotePr>
                    <w:pos w:val="docEnd"/>
                    <w:numFmt w:val="upperLetter"/>
                    <w:numStart w:val="2"/>
                  </w:endnotePr>
                  <w:compat>
                    <w:useFELayout/>
                    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
                  </w:compat>
                </w:settings>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Settings</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal("doNotCompress", document.Settings.CharacterSpacingControlValue ?? string.Empty);
        TestAssert.Equal("720", document.Settings.DefaultTabStopValue ?? string.Empty);
        TestAssert.Equal(36d, document.Settings.DefaultTabStopPoints ?? 0d);
        TestAssert.True(document.Settings.UseFELayout == true, "Empty useFELayout should opt in.");
        TestAssert.Equal("1", document.Settings.RevisionViewSettings.MarkupValue ?? string.Empty);
        TestAssert.True(document.Settings.RevisionViewSettings.ShowMarkup == true, "revisionView markup=1 should opt in.");
        TestAssert.Equal("0", document.Settings.RevisionViewSettings.CommentsValue ?? string.Empty);
        TestAssert.True(document.Settings.RevisionViewSettings.ShowComments == false, "revisionView comments=0 should opt out.");
        TestAssert.True(document.Settings.RevisionViewSettings.ShowInsertionsAndDeletions == true, "revisionView insDel=1 should opt in.");
        TestAssert.True(document.Settings.RevisionViewSettings.ShowFormatting == false, "revisionView formatting=0 should opt out.");
        TestAssert.True(document.Settings.RevisionViewSettings.ShowInkAnnotations == true, "revisionView inkAnnotations=1 should opt in.");
        TestAssert.True(document.Settings.TrackChangesSettings.TrackRevisions == true, "Empty trackRevisions should opt in.");
        TestAssert.Equal("0", document.Settings.TrackChangesSettings.DoNotTrackMovesValue ?? string.Empty);
        TestAssert.True(document.Settings.TrackChangesSettings.DoNotTrackMoves == false, "doNotTrackMoves val=0 should opt out.");
        TestAssert.Equal("1", document.Settings.TrackChangesSettings.DoNotTrackFormattingValue ?? string.Empty);
        TestAssert.True(document.Settings.TrackChangesSettings.DoNotTrackFormatting == true, "doNotTrackFormatting val=1 should opt in.");
        TestAssert.True(document.Settings.MirrorMargins == true, "Empty mirrorMargins should opt in.");
        TestAssert.True(document.Settings.MirrorMarginsValue is null, "Val-less mirrorMargins should keep a null source token.");
        TestAssert.Equal("pageBottom", document.Settings.FootnoteReferenceSettings.PositionValue ?? string.Empty);
        TestAssert.Equal("lowerRoman", document.Settings.FootnoteReferenceSettings.NumberFormatValue ?? string.Empty);
        TestAssert.Equal(4, document.Settings.FootnoteReferenceSettings.NumberStart ?? 0);
        TestAssert.Equal("continuous", document.Settings.FootnoteReferenceSettings.NumberRestartValue ?? string.Empty);
        TestAssert.Equal("docEnd", document.Settings.EndnoteReferenceSettings.PositionValue ?? string.Empty);
        TestAssert.Equal("upperLetter", document.Settings.EndnoteReferenceSettings.NumberFormatValue ?? string.Empty);
        TestAssert.Equal(2, document.Settings.EndnoteReferenceSettings.NumberStart ?? 0);
        DocxCompatSetting compat = document.Settings.CompatSettings.Single();
        TestAssert.Equal("compatibilityMode", compat.Name ?? string.Empty);
        TestAssert.Equal("15", compat.Value ?? string.Empty);
    }

    public static void DocxReaderAppliesAutoLineSpacingAsTwoHundredFortieths()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr>
                      <w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>Auto line spacing</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal(1.15d, paragraph.LineSpacingFactor);
    }

    public static void DocxReaderUsesWordDefaultAutoLineSpacingWhenLineTokenIsAbsent()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:spacing w:before="36" w:after="0"/></w:pPr>
                      <w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Missing line token</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>No spacing token</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(2, document.Paragraphs.Count);
        TestAssert.Equal(1.2d, document.Paragraphs[0].LineSpacingFactor);
        TestAssert.Equal(1.2d, document.Paragraphs[1].LineSpacingFactor);
    }

    public static void DocxReaderPreservesRunFontTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:rPr>
                          <w:rFonts w:ascii="Corporate Sans" w:hAnsi="Corporate Sans" w:eastAsia="Yu Gothic" w:cs="Arial"
                            w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:csTheme="minorBidi"/>
                        </w:rPr>
                        <w:t>Font tokens</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxRunFonts fonts = document.Paragraphs.Single().Runs.Single().Fonts;
        TestAssert.Equal("Corporate Sans", fonts.Ascii ?? string.Empty);
        TestAssert.Equal("Corporate Sans", fonts.HighAnsi ?? string.Empty);
        TestAssert.Equal("Yu Gothic", fonts.EastAsia ?? string.Empty);
        TestAssert.Equal("Arial", fonts.ComplexScript ?? string.Empty);
        TestAssert.Equal("minorHAnsi", fonts.AsciiTheme ?? string.Empty);
        TestAssert.Equal("minorHAnsi", fonts.HighAnsiTheme ?? string.Empty);
        TestAssert.Equal("minorEastAsia", fonts.EastAsiaTheme ?? string.Empty);
        TestAssert.Equal("minorBidi", fonts.ComplexScriptTheme ?? string.Empty);
    }

    public static void DocxFontResolverBuildsLatinTypefaceCandidatesFromCatalogAndTheme()
    {
        var catalog = new DocxFontCatalog(
            [new DocxFontTableEntry("Corporate Sans", "Aptos", "swiss", "variable", null, null)],
            new DocxThemeFonts("Aptos Display", "Aptos", null, null, null, null));
        var run = new DocxTextRun("Text", 11d, null, false, false, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Corporate Sans",
                HighAnsi: null,
                EastAsia: null,
                ComplexScript: null,
                AsciiTheme: "minorHAnsi",
                HighAnsiTheme: null,
                EastAsiaTheme: null,
                ComplexScriptTheme: null)
        };

        DocxTypefaceCandidates candidates = DocxFontResolver.ResolveLatinTypeface(run, catalog);

        TestAssert.Equal("Corporate Sans", candidates.Primary ?? string.Empty);
        TestAssert.Equal("Aptos", candidates.Alternate ?? string.Empty);
        TestAssert.Equal("Aptos", candidates.Theme ?? string.Empty);
    }

    public static void DocxFontResolverBuildsEastAsianTypefaceCandidatesFromCatalogAndTheme()
    {
        var catalog = new DocxFontCatalog(
            [new DocxFontTableEntry("Corporate East", "Installed East", "modern", "variable", null, "80")],
            new DocxThemeFonts(
                "Theme Display",
                "Theme Sans",
                MajorComplexScriptTypeface: null, MinorComplexScriptTypeface: null, MajorEastAsiaTypeface: "Theme East Display",
                MinorEastAsiaTypeface: "Theme East"));
        var run = new DocxTextRun("\u6f22\u5b57", 11d, null, false, false, false, null, "Latin Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Latin Sans",
                HighAnsi: null,
                EastAsia: "Corporate East",
                ComplexScript: null,
                AsciiTheme: "majorHAnsi",
                HighAnsiTheme: null,
                EastAsiaTheme: "minorEastAsia",
                ComplexScriptTheme: null)
        };

        DocxTypefaceCandidates candidates = DocxFontResolver.ResolveLatinTypeface(run, catalog);

        TestAssert.Equal("Corporate East", candidates.Primary ?? string.Empty);
        TestAssert.Equal("Installed East", candidates.Alternate ?? string.Empty);
        TestAssert.Equal("Theme East", candidates.Theme ?? string.Empty);
    }

    public static void DocxFontPlanResolvesEastAsianAndComplexScriptThemeTypefaces()
    {
        var eastAsianRun = new DocxTextRun("\u6f22\u5b57", 11d, null, false, false, false, null, "Latin Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Latin Sans",
                HighAnsi: null,
                EastAsia: "Corporate East",
                ComplexScript: null,
                AsciiTheme: null,
                HighAnsiTheme: null,
                EastAsiaTheme: "minorEastAsia",
                ComplexScriptTheme: null)
        };
        var complexScriptRun = new DocxTextRun("\u0633\u0644\u0627\u0645", 11d, null, false, false, false, null, "Latin Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Latin Sans",
                HighAnsi: null,
                EastAsia: null,
                ComplexScript: "Corporate Bidi",
                AsciiTheme: null,
                HighAnsiTheme: null,
                EastAsiaTheme: null,
                ComplexScriptTheme: "minorBidi")
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(
            [eastAsianRun, complexScriptRun],
            new DocxFontCatalog(
                [
                    new DocxFontTableEntry("Corporate East", "Installed East", "modern", null, null, "80"),
                    new DocxFontTableEntry("Corporate Bidi", "Installed Bidi", "roman", null, null, "B2")
                ],
                new DocxThemeFonts(
                    "Theme Display",
                    "Theme Sans",
                    MajorComplexScriptTypeface: null, MajorEastAsiaTypeface: null, MinorComplexScriptTypeface: "Theme Bidi",
                    MinorEastAsiaTypeface: "Theme East")));
        var resolver = new MapFontResolver(["Theme East", "Theme Bidi"], "Resolver Fallback");
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);

        DocxResolvedRunTypeface eastAsian = fontPlan.Runs.Single(run => run.Run.Text == "\u6f22\u5b57");
        DocxResolvedRunTypeface complexScript = fontPlan.Runs.Single(run => run.Run.Text == "\u0633\u0644\u0627\u0645");

        TestAssert.Equal(DocxTypefaceResolutionSource.Theme, eastAsian.Source);
        TestAssert.Equal("Theme East", eastAsian.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Corporate East|Installed East|Theme East", string.Join("|", eastAsian.CandidateFamilies));
        TestAssert.Equal(DocxTypefaceResolutionSource.Theme, complexScript.Source);
        TestAssert.Equal("Theme Bidi", complexScript.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Corporate Bidi|Installed Bidi|Theme Bidi", string.Join("|", complexScript.CandidateFamilies));
    }

    public static void DocxFontPlanKeepsPrimaryBeforeAlternateAndTheme()
    {
        var run = new DocxTextRun("Text", 11d, null, true, true, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Corporate Sans",
                HighAnsi: null,
                EastAsia: null,
                ComplexScript: null,
                AsciiTheme: "majorHAnsi",
                HighAnsiTheme: null,
                EastAsiaTheme: null,
                ComplexScriptTheme: null)
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(
            run,
            new DocxFontCatalog(
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null, null)],
                new DocxThemeFonts("Theme Display", "Theme Sans", null, null, null, null)));
        var resolver = new MapFontResolver(["Corporate Sans", "Installed Sans", "Theme Display"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Single();

        TestAssert.Equal(DocxTypefaceResolutionSource.Primary, resolved.Source);
        TestAssert.Equal("Corporate Sans", resolved.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Corporate Sans", resolved.ResolvedFamily ?? string.Empty);
        TestAssert.True(resolved.Resolution?.Bold == true && resolved.Resolution?.Italic == true, "Expected run style to flow into the font request.");
    }

    public static void DocxFontPlanPrefersOfficeBodyFontForImplicitDefaultTypeface()
    {
        var run = new DocxTextRun("Office default", 11d, null, true, false, false, null, null);
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, DocxFontCatalog.Empty);
        var resolver = new MapFontResolver(["Aptos", "Calibri", "Arial"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Single();

        TestAssert.Equal(DocxTypefaceResolutionSource.ResolverFallback, resolved.Source);
        TestAssert.Equal(DocxRenderer.DefaultDocumentTypefaceRequest, resolved.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Aptos", resolved.ResolvedFamily ?? string.Empty);
        TestAssert.True(resolved.Resolution?.Bold == true, "Implicit default runs should preserve requested style when resolving the Office body fallback.");
    }

    public static void DocxFontPlanUsesBodyElementInventoryAsCanonicalSource()
    {
        var bodyRun = new DocxTextRun("Body", 11d, null, false, false, false, null, "Body Sans")
        {
            Fonts = new DocxRunFonts("Body Sans", null, null, null, null, null, null, null)
        };
        DocxParagraph bodyParagraph = DocxTests.CreateFontPlanParagraph(bodyRun);
        var cellRun = new DocxTextRun("Cell", 11d, null, false, false, false, null, "Cell Sans")
        {
            Fonts = new DocxRunFonts("Cell Sans", null, null, null, null, null, null, null)
        };
        DocxParagraph cellParagraph = DocxTests.CreateFontPlanParagraph(cellRun);
        var table = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [cellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], null)]);
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph), new DocxTableElement(table)],
            [],
            []);
        var resolver = new MapFontResolver(["Body Sans", "Cell Sans"], "Resolver Fallback");

        string plannedTexts = string.Join("|", DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Select(run => run.Run.Text).Order(StringComparer.Ordinal));

        TestAssert.Equal("Body|Cell", plannedTexts);
    }

    public static void DocxBlockTraversalNormalizesStaticStoryBodyFallbacks()
    {
        DocxParagraph legacyFallback = DocxTests.CreateDocxLayoutParagraph("Legacy", 12d, 12d);
        DocxParagraph paragraphFallback = DocxTests.CreateDocxLayoutParagraph("ParagraphMap", 12d, 12d);
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("BodyMap", 12d, 12d);
        var bodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
        {
            ["default"] = [new DocxParagraphElement(bodyParagraph)]
        };
        var paragraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
        {
            ["default"] = [paragraphFallback]
        };

        IReadOnlyList<DocxParagraph> referenced = DocxBlockTraversal
            .EnumerateReferencedStaticStoryParagraphs(bodyElementsByType, paragraphsByType, [legacyFallback])
            .ToArray();
        bool foundBodyElements = DocxBlockTraversal.TryGetStaticStoryBodyElements(
            "default",
            bodyElementsByType,
            paragraphsByType,
            out IReadOnlyList<DocxBodyElement> selectedBodyElements);
        bool foundFallbackElements = DocxBlockTraversal.TryGetStaticStoryBodyElements(
            "default",
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase),
            paragraphsByType,
            out IReadOnlyList<DocxBodyElement> selectedFallbackElements);
        IReadOnlyList<DocxParagraph> legacyOnly = DocxBlockTraversal
            .EnumerateReferencedStaticStoryParagraphs(
                new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase),
                new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase),
                [legacyFallback])
            .ToArray();

        TestAssert.Equal("BodyMap", referenced.Single().Runs.Single().Text);
        TestAssert.True(foundBodyElements, "Expected canonical static body elements to be selectable.");
        TestAssert.Equal("BodyMap", ((DocxParagraphElement)selectedBodyElements.Single()).Paragraph.Runs.Single().Text);
        TestAssert.True(foundFallbackElements, "Expected paragraph-map compatibility views to be selectable when body elements are absent.");
        TestAssert.Equal("ParagraphMap", ((DocxParagraphElement)selectedFallbackElements.Single()).Paragraph.Runs.Single().Text);
        TestAssert.Equal("Legacy", legacyOnly.Single().Runs.Single().Text);
    }

    public static void DocxDocumentCompatibilityInventoriesDeriveFromBodyElements()
    {
        DocxParagraph fallbackParagraph = DocxTests.CreateDocxLayoutParagraph("Fallback", 12d, 12d);
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body", 12d, 12d);
        DocxParagraph nestedParagraph = DocxTests.CreateDocxLayoutParagraph("Nested", 12d, 12d);
        var nestedTable = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxTableCell cell = new(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var bodyTable = new DocxTable(null, [60d], [new DocxTableRow([cell], 30d)]);
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph), new DocxTableElement(bodyTable)],
            [fallbackParagraph],
            []);

        TestAssert.Equal("Body", document.Paragraphs.Single().Runs.Single().Text);
        TestAssert.Equal("Body|Nested", string.Join("|", DocxBlockTraversal.EnumerateBodyParagraphs(document).Select(paragraph => paragraph.Runs.Single().Text)));
        TestAssert.Equal("Body", string.Join("|", DocxBlockTraversal.EnumerateDirectParagraphs(document.BodyElements).Select(paragraph => paragraph.Runs.Single().Text)));
        TestAssert.Equal(2, document.Tables.Count);
    }

    public static void DocxRelatedStoryCompatibilityInventoriesDeriveFromBodyElements()
    {
        DocxParagraph fallbackParagraph = DocxTests.CreateDocxLayoutParagraph("Fallback", 12d, 12d);
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Story", 12d, 12d);
        DocxParagraph nestedParagraph = DocxTests.CreateDocxLayoutParagraph("Nested", 12d, 12d);
        var nestedTable = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxTableCell cell = new(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var storyTable = new DocxTable(null, [60d], [new DocxTableRow([cell], 30d)]);
        var story = new DocxRelatedStory(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(bodyParagraph), new DocxTableElement(storyTable)],
            [fallbackParagraph],
            [], null);

        TestAssert.Equal("Story", story.Paragraphs.Single().Runs.Single().Text);
        TestAssert.Equal("Story|Nested", string.Join("|", DocxBlockTraversal.EnumerateBodyParagraphs(story).Select(paragraph => paragraph.Runs.Single().Text)));
        TestAssert.Equal("Story", string.Join("|", DocxBlockTraversal.EnumerateDirectParagraphs(story.BodyElements).Select(paragraph => paragraph.Runs.Single().Text)));
        TestAssert.Equal(2, story.Tables.Count);
    }

    public static void DocxFontPlanSnapshotReportsPrivateSafeCounts()
    {
        var primaryRun = new DocxTextRun("Primary", 11d, null, false, false, false, null, "Primary Sans")
        {
            Fonts = new DocxRunFonts("Primary Sans", null, null, null, null, null, null, null)
        };
        var alternateRun = new DocxTextRun("Alternate", 11d, null, false, false, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts("Corporate Sans", null, null, null, null, null, null, null)
        };
        var missingRun = new DocxTextRun("Missing", 11d, null, false, false, false, null, null);
        DocxDocument document = DocxTests.CreateFontPlanDocument(
            [primaryRun, alternateRun, missingRun],
            new DocxFontCatalog(
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null, null)],
                DocxThemeFonts.Empty));
        var resolver = new MapFontResolver(["Primary Sans", "Installed Sans"], "Resolver Fallback");

        DocxFontPlanSnapshot snapshot = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectFontPlan(document);

        TestAssert.Equal(3, snapshot.RunCount);
        TestAssert.Equal(1, snapshot.PrimaryCount);
        TestAssert.Equal(1, snapshot.FontTableAlternateCount);
        TestAssert.Equal(0, snapshot.ThemeCount);
        TestAssert.Equal(1, snapshot.ResolverFallbackCount);
        TestAssert.Equal(0, snapshot.MissingCount);
        TestAssert.Equal(3, snapshot.DistinctCandidateFamilyCount);
        TestAssert.Equal(3, snapshot.DistinctResolvedFamilyCount);
    }

    public static void DocxFontPlanSnapshotReportsPrivateSafeOpenTypeMetrics()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        var run = new DocxTextRun("Metric probe", 10d, null, false, false, false, null, font.Value.Resolution.FamilyName)
        {
            Fonts = new DocxRunFonts(font.Value.Resolution.FamilyName, null, null, null, null, null, null, null)
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));

        DocxFontPlanSnapshot snapshot = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectFontPlan(document);

        DocxFontMetricBucketSnapshot bucket = snapshot.MetricBuckets.Single();
        TestAssert.Equal(DocxTypefaceResolutionSource.Primary.ToString(), bucket.Source);
        TestAssert.Equal(10d, bucket.FontSize);
        TestAssert.Equal(1, bucket.RunCount);
        TestAssert.Equal(12, bucket.ResolvedFamilyHash?.Length ?? 0);
        TestAssert.Equal(font.Value.Font.UnitsPerEm, bucket.UnitsPerEm ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.TypographicAscender, bucket.TypographicAscender ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.TypographicDescender, bucket.TypographicDescender ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.TypographicLineGap, bucket.TypographicLineGap ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.WindowsAscender, bucket.WindowsAscender ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.WindowsDescender, bucket.WindowsDescender ?? 0);
        TestAssert.Equal(DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font.Value.Font, 10d), bucket.SingleLineHeightPoints ?? 0d);
        TestAssert.Equal(DocxLineMetrics.MeasureWindowsAscender(font.Value.Font, 10d), bucket.WindowsAscenderPoints ?? 0d);
        TestAssert.Equal(DocxLineMetrics.MeasureWindowsDescender(font.Value.Font, 10d), bucket.WindowsDescenderPoints ?? 0d);
    }
}
