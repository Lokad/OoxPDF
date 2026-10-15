using System.Buffers.Binary;
using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathFractionTests
{
    public static void MathFontRadicalDegreesUseRoleSizingKernsAndInkExtents()
    {
        byte[] bytes=RadicalFontBytes();
        var range=TestFontBuilder.GetTableRange(bytes,"MATH");
        byte[] math=bytes.AsSpan(range.Offset,range.Length).ToArray();
        BinaryPrimitives.WriteInt16BigEndian(math.AsSpan(18+(int)OpenTypeMathConstant.RadicalKernBeforeDegree*4),100);
        BinaryPrimitives.WriteInt16BigEndian(math.AsSpan(18+(int)OpenTypeMathConstant.RadicalKernAfterDegree*4),-250);
        bytes=TestFontBuilder.ReplaceTable(bytes,"MATH",math);
        OpenTypeFont font=OpenTypeFont.Load(bytes);
        var resource=new PdfFontResource("Degree",PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[]{(font.MapCodePoint('B'),(int)'B'),(font.MapCodePoint('+'),(int)'+'),(font.MapCodePoint('C'),(int)'C'),((ushort)130,0x221A),((ushort)128,0x221A)},CancellationToken.None));
        DocxEquation source=RadicalEquation(12d);
        var root=(DocxMathRadical)source.Content.Children[0];
        foreach(bool? visible in new bool?[]{false,null})
        {
            var radical=root with { DegreeHidden=visible,Degree=Row("B+C") };
            var equation=source with { Content=new([radical]) };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation,CancellationToken.None,out var runs) &&
                runs.Count==2 && runs[0].ScriptLevel==0 && runs[0].RadicalOwner==radical && runs[1].ScriptLevel==2,
                "The degree has script-script role sizing while the radicand retains the implicit root owner.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation,_=>resource,CancellationToken.None,out var box,out var reason),
                reason??"A literal visible degree must compose with the owned root.");
            DocxMathBox measured=TestAssert.NotNull(box);
            var degree=measured.Glyphs.Skip(4).ToArray();
            TestAssert.True(degree.Length==3 && degree.All(g=>g.FontSize==6d),
                "The synthetic font's50% script-script constant owns degree size independently of its73% script scale.");
            TestAssert.True(degree.All(g=>g.BaselineY>measured.Glyphs[0].BaselineY) &&
                measured.AboveBaseline>=degree.Max(g=>g.BaselineY+700d*g.FontSize/1000d),
                "Raised degree ink is included in the composed extent.");
            TestAssert.True(measured.Glyphs[0].X>=degree[0].X && measured.Rules[0].X==measured.Glyphs[1].X,
                "Degree kerns shift root, radicand and overbar together.");
            TestAssert.True(Math.Abs(degree[0].X-1.2d)<1e-9 &&
                Math.Abs(measured.Glyphs[0].X-(degree[0].X+degree.Sum(g=>g.Advance)-3d))<1e-9,
                "Signed kerns use parent design units, retaining their overlap rather than child-size scaling.");
            var larger=radical with { Degree=radical.Degree with { ArgumentProperties=new(1,false) } };
            TestAssert.True(DocxMathLayout.TryGetRuns(source with { Content=new([larger]) },CancellationToken.None,out var adjusted) &&
                adjusted[1].ScriptLevel==1,"Signed argument sizing adjusts the role's starting level through the shared rule.");
        }
    }

    public static void MathFontRadicalAuthoredControlsKeepPerNodeFontsSizesAndChildMetrics()
    {
        var small=RadicalResource();var large=RadicalResource();
        var original=(DocxMathRadical)RadicalEquation(12d).Content.Children[0];
        var a=original with { ControlText=new("",8d,"008000",false,false,false,null,"TestFont") { Italic=true },ControlFontFamily="Small" };
        var b=original with { ControlText=new("",20d,"000000",false,false,false,null,"TestFont"),ControlFontFamily="Large" };
        var source=RadicalEquation(12d) with { Content=new([a,Row("+"),b]) };
        TestAssert.True(DocxMathLayout.TryMeasure(source,_=>small,CancellationToken.None,out var box,out var reason,
            resolveRadicalFont:root=>ReferenceEquals(root,a)?small:large),reason??"Independent control nodes own their own implicit font resources.");
        DocxMathBox measured=TestAssert.NotNull(box);var roots=measured.Glyphs.Where(g=>g.SourceCodePoint==0x221A).ToArray();
        TestAssert.True(roots.Length==2 && roots[0].Font==small && roots[1].Font==large &&
            Math.Abs(roots[0].FontSize-8.04d)<1e-9 && Math.Abs(roots[1].FontSize-20.04d)<1e-9,
            "Each authored control keeps its glyph resource and export size; no first-root ownership leaks to siblings.");
        TestAssert.True(roots.All(g=>g.TextShear==0d) && measured.Glyphs.Where(g=>g.SourceCodePoint!=0x221A).All(g=>g.FontSize==12d),
            "Italic control flags do not shear mathematical root symbols or rescale authored children.");
        TestAssert.True(!DocxMathLayout.TryMeasure(source,_=>small,CancellationToken.None,out var missing,out _,
            resolveRadicalFont:root=>ReferenceEquals(root,a)?small:null) && missing is null,
            "A missing explicit glyph owner cannot silently adopt a sibling or first child resource.");
        foreach(var control in new[]{a.ControlText! with { Underline=true },a.ControlText! with { Hidden=true },a.ControlText! with { Bold=true }})
            TestAssert.True(!DocxMathLayout.TryMeasure(source with { Content=new([a with { ControlText=control }]) },_=>small,
                CancellationToken.None,out _,out _),"Unqualified decorative/hidden/bold controls retain the shared formatting fallback.");
    }

    public static void MathFontRadicalNormalChildrenKeepIndependentParentGlyphsAndControlColor()
    {
        var parent=RadicalResource();
        var text=new PdfFontResource("Text",PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),"B+C".Select(c=>(int)c),CancellationToken.None));
        var child=new DocxMathRun(new("B+C",20d,"FF0000",false,false,false,null,"TestFont"),null,null) { NormalTextValue=true };
        var radical=new DocxMathRadical(new([child]),new([]),true)
            { ControlText=new("",12d,"000000",false,false,false,null,"TestFont") };
        var source=new DocxEquation(new([radical]),0,0,true,null,null,null) { ParentFontSize=12d };
        TestAssert.True(!DocxMathLayout.TryMeasure(source,_=>text,CancellationToken.None,out _,out _,parent.Font.Font),
            "Metrics alone cannot encode a parent root glyph in a child text-face resource.");
        TestAssert.True(DocxMathLayout.TryMeasure(source,_=>text,CancellationToken.None,out var box,out var reason,parent.Font.Font,_=>parent),
            reason??"A glyph-bearing parent owns the root independently of the normal-text face.");
        DocxMathBox measured=TestAssert.NotNull(box);
        TestAssert.True(measured.Glyphs[0].Font==parent && measured.Glyphs.Skip(1).All(g=>g.Font==text),
            "Root and radicand retain their independently prepared physical font resources.");
        TestAssert.True(measured.Glyphs[0].Color==default && measured.Rules[0].Color==default &&
            measured.Glyphs.Skip(1).All(g=>g.Color.Red==255 && g.Color.Green==0 && g.Color.Blue==0),
            "Inherited black control color cannot adopt the child's authored red text.");
        var graphics=new PdfGraphicsBuilder();DocxMathBoxPainter.Paint(measured,graphics,20d,50d,CancellationToken.None);
        using var output=new MemoryStream();
        PdfDocumentWriter.WriteBlank(output,[new PdfPage(200d,150d,graphics.ToString(),[parent,text])],CancellationToken.None);
        TestAssert.True(output.Length>0,"Both ownership lanes retain valid glyph encodings and resources.");
    }

    public static void MathFontRadicalBoxesKeepParentMetricsSelectedGlyphsAndOwnedRules()
    {
        var resource=RadicalResource();
        foreach(double childSize in new[]{8d,20d})
        {
            DocxEquation equation=RadicalEquation(childSize);
            TestAssert.True(DocxMathLayout.TryMeasure(equation,_=>resource,CancellationToken.None,out var box,out var reason),
                reason??"Expected a ready-made owned radical box.");
            DocxMathBox measured=TestAssert.NotNull(box);
            TestAssert.Equal(4,measured.Glyphs.Count);TestAssert.Equal(1,measured.Rules.Count);
            var root=measured.Glyphs[0];
            TestAssert.True(root.SourceCodePoint==0x221A && root.FontSize==12d && root.OriginalGlyph==(childSize==8d?130:128),
                "Parent metrics retain the12pt implicit root and select a larger owned glyph for taller child ink.");
            TestAssert.True(measured.Glyphs.Skip(1).All(g=>Math.Abs(g.FontSize-OfficePdfTextEmissionProfile.FontSize(childSize))<1e-9),
                "Radicand glyphs retain their independently authored export size.");
            double inkTop=measured.Glyphs.Skip(1).Max(g=>700d*g.FontSize/1000d);
            TestAssert.True(measured.Rules[0].Y-inkTop>=80d*12d/1000d-1e-9,
                "The overbar remains above all child ink by at least the parent's display gap.");
            TestAssert.True(DocxMathLayout.TryGetLineMetrics(measured,CancellationToken.None,out double above,out _) &&
                Math.Abs(above-measured.AboveBaseline-120d*12d/1000d)<1e-9,
                "An explicit parent owns one leading term even when child math glyphs are larger.");
            var graphics=new PdfGraphicsBuilder();DocxMathBoxPainter.Paint(measured,graphics,20d,50d,CancellationToken.None);
            using var output=new MemoryStream();
            PdfDocumentWriter.WriteBlank(output,[new PdfPage(200d,150d,graphics.ToString(),[resource])],CancellationToken.None);
            TestAssert.True(output.Length>0,"Root variants and the bar retain valid glyph/encoding/resource ownership.");
        }
    }

    public static void MathFontRadicalBoxesKeepDegreeAssemblyEncodingAndCancellationFallbacks()
    {
        var resource=RadicalResource();DocxEquation source=RadicalEquation(12d);
        var radical=(DocxMathRadical)source.Content.Children[0];
        foreach(var invalid in new[]{radical with { DegreeHidden=false },radical with { Degree=Row("B") },
            radical with { HasUnsupportedProperties=true },radical with { Radicand=new([source.Content.Children[0]]) }})
            TestAssert.True(!DocxMathLayout.TryMeasure(source with { Content=new([invalid]) },_=>resource,CancellationToken.None,out var box,out _) && box is null,
                "Unsupported degree/properties/deeper structures keep the complete radical fallback.");
        TestAssert.True(!DocxMathLayout.TryMeasure(RadicalEquation(100d),_=>resource,CancellationToken.None,out var large,out _) && large is null,
            "An assembly-only height cannot publish the largest inadequate ready-made glyph.");
        var missing=new PdfFontResource("Missing",PdfEmbeddedFont.Create(resource.Font.Font,"B+C".Select(c=>(int)c),CancellationToken.None));
        TestAssert.True(!DocxMathLayout.TryMeasure(source,_=>missing,CancellationToken.None,out var absent,out _) && absent is null,
            "A selected variant needs its own encoded CID before publishing the box.");
        using var cancelled=new CancellationTokenSource();cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(()=>DocxMathLayout.TryMeasure(source,_=>resource,cancelled.Token,out _,out _));
        using(OoxConversionBudget.Scope scope=OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxFontWorkPerConversion=0 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(()=>DocxMathLayout.TryMeasure(source,_=>resource,CancellationToken.None,out _,out _));
    }

    public static void MathFontRadicalAliasesKeepAuthoredUnicodeInEitherSourceOrder()
    {
        XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m="http://schemas.openxmlformats.org/officeDocument/2006/math";
        foreach(bool rootFirst in new[]{true,false})
        {
            string fixture=Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,"..","..","..","Cases","docx-omml-radical-display-large20.docx"));
            using ZipArchive archive=ZipFile.OpenRead(fixture);
            var parts=new Dictionary<string,byte[]>();
            foreach(var entry in archive.Entries){using var input=entry.Open();using var data=new MemoryStream();input.CopyTo(data);parts.Add(entry.FullName,data.ToArray());}
            XElement document=XElement.Parse(System.Text.Encoding.UTF8.GetString(parts["word/document.xml"]));
            XElement root=document.Descendants(m+"rad").Single();
            XElement run=root.Element(m+"e")!.Element(m+"r")!;run.Element(m+"t")!.Value="B+C";
            XElement props=run.Element(m+"rPr")??new XElement(m+"rPr");if(props.Parent is null)run.AddFirst(props);
            props.RemoveNodes();props.Add(new XElement(m+"sty",new XAttribute(m+"val","p")));
            XElement rootParagraph=root.Ancestors(w+"p").First();
            XElement alias=new(w+"p",new XElement(m+"oMathPara",new XElement(m+"oMath",new XElement(m+"r",
                new XElement(m+"rPr",new XElement(m+"sty",new XAttribute(m+"val","p"))),new XElement(m+"t","\u009F")))));
            if(rootFirst)rootParagraph.AddAfterSelf(alias);else rootParagraph.AddBeforeSelf(alias);
            parts["word/document.xml"]=System.Text.Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
            string path=TestFixtures.WriteTempPackage(".docx",parts);
            try
            {
                DocxDocument source=DocxTests.ReadDocx(path,OoxPdfDocxMarkupMode.Final);
                var renderer=new DocxRenderer(new RadicalFontResolver(RadicalFontBytes()),OoxPdfDocxMarkupMode.Final,OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                string text=string.Concat(renderer.RenderBlankPages(source,null,CancellationToken.None).Select(DocxPageFootnoteNumberingTests.DecodeText));
                TestAssert.True(text.Contains('\u009F'),"Authored Unicode must survive an implicit-root variant alias in either source order.");
                TestAssert.True(!text.Contains('\u221A'),"An unresolved glyph/Unicode alias keeps root fallback instead of changing the authored source mapping.");
            }
            finally{File.Delete(path);}
        }
    }

    private sealed class RadicalFontResolver(byte[] bytes):IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)=>new(request.FamilyName,"TestFont",new FontStyleKey(HasMathTable:true),
            new MemoryFontProgramSource("test:radical-alias",bytes),false);
    }

    private static DocxEquation RadicalEquation(double size)=>new(new([new DocxMathRadical(
        new([new DocxMathRun(new("B+C",size,null,false,false,false,null,"TestFont"),"p",null)]),new([]),true)]),0,0,true,null,null,null)
        { ParentFontSize=12d };

    private static PdfFontResource RadicalResource()
    {
        OpenTypeFont font=OpenTypeFont.Load(RadicalFontBytes());
        return new("Root",PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[]{(font.MapCodePoint('B'),(int)'B'),(font.MapCodePoint('+'),(int)'+'),(font.MapCodePoint('C'),(int)'C'),((ushort)130,0x221A),((ushort)128,0x221A)},CancellationToken.None));
    }

    private static byte[] RadicalFontBytes()
    {
        byte[] math=Constants();Array.Resize(ref math,254);
        void Set(int at,ushort value)=>BinaryPrimitives.WriteUInt16BigEndian(math.AsSpan(at),value);
        Set(12,50);Set(222,65);
        Set(8,224);Set(224+2,12);Set(224+6,1);Set(224+10,18);Set(224+12,1);Set(224+14,1);Set(224+16,130);
        Set(224+18+2,2);Set(224+18+4,130);Set(224+18+6,1000);Set(224+18+8,128);Set(224+18+10,2000);
        foreach((OpenTypeMathConstant field,ushort value) in new[]{(OpenTypeMathConstant.RadicalVerticalGap,(ushort)50),
            (OpenTypeMathConstant.RadicalDisplayStyleVerticalGap,(ushort)80),(OpenTypeMathConstant.RadicalRuleThickness,(ushort)40),
            (OpenTypeMathConstant.RadicalExtraAscender,(ushort)40)})Set(18+(int)field*4,value);
        byte[] cmap=new byte[52];
        void U16(int at,ushort value)=>BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(at),value);
        void U32(int at,uint value)=>BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(at),value);
        U16(2,1);U16(4,3);U16(6,10);U32(8,12);U16(12,12);U32(16,40);U32(24,2);
        U32(28,0x20);U32(32,0x9F);U32(36,1);U32(40,0x221A);U32(44,0x221A);U32(48,130);
        return TestFontBuilder.ReplaceTable(TestFontBuilder.CreateMathFont(math),"cmap",cmap);
    }

    public static void MathFontStructuredSequencesRetainIndependentChildrenAndRules()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation child = Fraction();
        var fraction = (DocxMathFraction)child.Content.Children[0];
        var scripts = new DocxMathScripts(Row("B"), null, Row("C"));
        var source = child with { Content = new([Row("A+"), fraction, Row("+"), scripts, Row("+"), fraction, Row("+B")]) };
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
            reason ?? "A supported sequence must not inherit the former single-structure guard.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal("A+A+BC+BC+A+BC+B", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        TestAssert.Equal(2, measured.Rules.Count);
        TestAssert.True(measured.Rules[0].X + measured.Rules[0].Width < measured.Rules[1].X,
            "Independent child rules retain distinct source positions in the sequence.");
        TestAssert.True(DocxMathLayout.TryMeasure(child with { Content = new([scripts]) }, _ => resource,
            CancellationToken.None, out var scriptBox, out _), "Expected the independently measured script.");
        TestAssert.True(measured.Glyphs.Single(g => g.SourceCodePoint == 'C' && g.FontSize < 12d).BaselineY ==
            TestAssert.NotNull(scriptBox).Glyphs[1].BaselineY,
            "The script role retains its independently measured style/baseline, including zero-shift synthetic metrics.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(300d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "All composed glyphs/rules retain valid owned PDF resources.");
    }

    public static void MathFontStructuredSequencesConsumeParentLeadingOnceForNormalChildren()
    {
        OpenTypeFont parent = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(Constants()));
        var text = new PdfFontResource("Text1", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),
            "A+BC".Select(c => (int)c), CancellationToken.None));
        DocxMathRow Normal(string value) => new([new DocxMathRun(new(value, 12d, null, false, false, false, null, "TestFont"),
            null, null) { NormalTextValue = true }]);
        var fraction = new DocxMathFraction(Normal("A+B"), Normal("C"), "bar");
        var one = new DocxEquation(new([fraction]), 0, 0, true, null, null, null) { ParentFontSize = 12d };
        var source = one with { Content = new([fraction, Normal("+"), fraction]) };
        TestAssert.True(DocxMathLayout.TryMeasure(one, _ => text, CancellationToken.None, out var child, out _, parent), "Expected existing normal-child fraction.");
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => text, CancellationToken.None, out var box, out _, parent), "Expected normal-child sequence.");
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(TestAssert.NotNull(child), CancellationToken.None, out double expectedAbove, out double expectedHeight) &&
            DocxMathLayout.TryGetLineMetrics(TestAssert.NotNull(box), CancellationToken.None, out double actualAbove, out double actualHeight) &&
            Math.Abs(actualAbove - expectedAbove) < 1e-9 && Math.Abs(actualHeight - expectedHeight) < 1e-9,
            "Side-by-side boxes consume one parent-leading/occupied line frame, not one term per fraction.");
        TestAssert.True(box!.ParentMetrics is { } owner && ReferenceEquals(owner.Font, parent) && box.Glyphs.All(g => g.Font == text),
            "The hidden parent metric owner survives independently of actual text-face glyph resources.");
    }

    public static void MathFontScriptRowsUseEquationSpacingWithoutChangingChildMetrics()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxMathRow Sized(string text) => new([new DocxMathRun(new(text, 20d, null, false, false, false, null, "TestFont"), "p", null)]);
        var scripts = new DocxMathScripts(Sized("B"), null, Sized("C"));
        DocxEquation child = Fraction() with { Content = new([scripts]), ParentFontSize = 12d };
        DocxMathRow prefix = Row("A+");
        TestAssert.True(DocxMathRowLayout.TryMeasure(prefix, _ => resource, CancellationToken.None, out var first, out _), "Expected literal prefix.");
        TestAssert.True(DocxMathLayout.TryMeasure(child, _ => resource, CancellationToken.None, out var basis, out _), "Expected authored20pt child.");
        TestAssert.True(DocxMathLayout.TryMeasure(child with { Content = new([prefix, scripts]) }, _ => resource,
            CancellationToken.None, out var box, out _), "Expected mixed-size structured row.");
        DocxMathBox measured = TestAssert.NotNull(box), standalone = TestAssert.NotNull(basis);
        TestAssert.True(Math.Abs(measured.Glyphs[2].X - TestAssert.NotNull(first).AdvanceWidth - 8d * 12d / 18d) < 1e-9,
            "Both sides of the binary boundary use the equation's12pt spacing, independently of20pt child glyphs.");
        TestAssert.True(measured.Glyphs[2].FontSize == standalone.Glyphs[0].FontSize &&
            measured.Glyphs[3].FontSize == standalone.Glyphs[1].FontSize &&
            measured.Glyphs[3].BaselineY == standalone.Glyphs[1].BaselineY,
            "Equation glue cannot rescale or reposition the already measured script role internally.");
    }

    public static void MathFontScriptRowsRetainChildStylesBaselinesAndSourceOwners()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        foreach (bool superscript in new[] { true, false })
        {
            var scripts = new DocxMathScripts(Row("B"), superscript ? null : Row("C"), superscript ? Row("C") : null);
            DocxEquation child = Fraction() with { Content = new([scripts]) };
            DocxEquation source = child with { Content = new([Row("A+"), scripts, Row("+B")]) };
            TestAssert.True(DocxMathLayout.TryGetRuns(source, CancellationToken.None, out var runs) &&
                runs.Count == 4 && runs[0].ScriptLevel == 0 && runs[1].ScriptLevel == 0 &&
                runs[2].ScriptLevel == 1 && runs[3].ScriptLevel == 0,
                "Literal siblings and base retain text style while the script role retains script style.");
            TestAssert.True(DocxMathLayout.TryMeasure(child, _ => resource, CancellationToken.None, out var basis, out _), "Expected existing script box.");
            TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
                reason ?? "One supported script must compose alongside literal atoms.");
            DocxMathBox measured = TestAssert.NotNull(box), scriptBox = TestAssert.NotNull(basis);
            TestAssert.Equal("A+BC+B", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            for (int i = 0; i < scriptBox.Glyphs.Count; i++)
            {
                var expected = scriptBox.Glyphs[i]; var actual = measured.Glyphs[i + 2];
                TestAssert.True(actual.Font == expected.Font && actual.FontSize == expected.FontSize &&
                    actual.OriginalGlyph == expected.OriginalGlyph && actual.BaselineY == expected.BaselineY,
                    "Composition cannot change the already qualified child's glyph/style/baseline ownership.");
            }
            TestAssert.True(measured.AdvanceWidth > scriptBox.AdvanceWidth && measured.Width >= measured.AdvanceWidth,
                "Literal siblings extend logical advance while occupied ink remains bounded by the stored frame.");
            var graphics = new PdfGraphicsBuilder();
            DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "The composed row retains valid source encodings/resources.");
        }
    }

    public static void MathFontScriptRowsKeepUnknownDeeperFallbackAndCancellation()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        var scripts = new DocxMathScripts(Row("B"), null, Row("C"));
        DocxEquation source = Fraction() with { Content = new([Row("A+"), scripts]) };
        foreach (var content in new[] { new DocxMathRow([Row("A+"), scripts with { PropertiesXml = ["<m:sSupPr/>"] }]),
            new DocxMathRow([Row("A+"), scripts with { Base = new([scripts]) }]) })
            TestAssert.True(!DocxMathLayout.TryGetRuns(source with { Content = content }, CancellationToken.None, out _) &&
                !DocxMathLayout.TryMeasure(source with { Content = content }, _ => resource, CancellationToken.None, out var box, out _) && box is null,
                "Unknown/deeper structures retain the complete source fallback.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => resource, cancelled.Token, out _, out _));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontFractionRowsShareBinaryBoundarySpacingAndRetainSourceOrder()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation child = Fraction();
        var fraction = (DocxMathFraction)child.Content.Children[0];
        DocxMathRow before = Row("A+"), after = Row("+B");
        var source = child with { Content = new([before, fraction, after]), ParentFontSize = 12d };
        TestAssert.True(DocxMathLayout.TryMeasure(child, _ => resource, CancellationToken.None, out var basis, out _), "Expected existing fraction.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(before, _ => resource, CancellationToken.None, out var first, out _), "Expected standalone prefix.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(after, _ => resource, CancellationToken.None, out var last, out _), "Expected standalone suffix.");
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
            reason ?? "Existing source atoms and one fraction must form a shared row.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal("A+A+BC+B", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        DocxMathBox firstBox = TestAssert.NotNull(first), fractionBox = TestAssert.NotNull(basis), lastBox = TestAssert.NotNull(last);
        double standalone = firstBox.AdvanceWidth + fractionBox.AdvanceWidth + lastBox.AdvanceWidth;
        TestAssert.True(Math.Abs(measured.AdvanceWidth - standalone - 16d * 12d / 18d) < 1e-9,
            "A binary operator around an ordinary fraction receives the shared four-unit glue on both sides, without unary boundary reclassification.");
        TestAssert.Equal(1, measured.Rules.Count);
        TestAssert.True(measured.Rules[0].X > firstBox.AdvanceWidth && measured.Rules[0].Width == fractionBox.Rules[0].Width,
            "The child rule keeps its width while moving with its source-owned row position.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Shared row encodings/resources must remain valid.");
    }

    public static void MathFontFractionRowsKeepDeeperStructureFallbackCancellationAndWorkBounds()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        source = source with { Content = new([Row("A+"), fraction]) };
        var deeper = source with { Content = new([Row("A+"), fraction with { Numerator = new([fraction]) }]) };
        TestAssert.True(!DocxMathLayout.TryGetRuns(deeper, CancellationToken.None, out _) &&
            !DocxMathLayout.TryMeasure(deeper, _ => resource, CancellationToken.None, out var box, out _) && box is null,
            "Unsupported deeper structures cannot publish partial formula text.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => resource, cancelled.Token, out _, out _));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontFractionScriptBasesRetainChildRulesRolesAndSelectedStyles()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        foreach (bool display in new[] { true, false })
        foreach (string type in new[] { "bar", "noBar" })
        {
            var scripts = new DocxMathScripts(new([fraction with { TypeValue = type }]), Row("B"), Row("B"));
            source = source with { Display = display, Content = new([scripts]) };
            TestAssert.True(DocxMathLayout.TryGetRuns(source, CancellationToken.None, out var runs), "Expected typed nested-base discovery.");
            TestAssert.True(runs.Count == 4 && runs[0].ScriptLevel == (display ? 0 : 1) && runs[2].ScriptLevel == 1,
                "Base fraction arguments inherit the equation style while script arguments own script style.");
            TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
                reason ?? "The existing boxes must compose without dropping the base.");
            DocxMathBox measured = TestAssert.NotNull(box);
            TestAssert.Equal("A+BCBB", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            TestAssert.Equal(type == "bar" ? 1 : 0, measured.Rules.Count);
            TestAssert.True(measured.Glyphs[^2].BaselineY <= measured.Glyphs[3].BaselineY &&
                measured.Glyphs[^1].BaselineY >= measured.Glyphs[0].BaselineY &&
                measured.Glyphs[^2].BaselineY < measured.Glyphs[^1].BaselineY,
                "Nested-base scripts retain their lower/upper roles; zero minimum-drop constants permit equal child baselines.");
            var graphics = new PdfGraphicsBuilder();
            DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "Nested glyphs and rules must retain valid owned resources/encodings.");
        }
    }

    public static void MathFontFractionScriptBasesRetainDeeperFallbackCancellationAndWorkBounds()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        var scripts = new DocxMathScripts(new([fraction]), null, Row("B"));
        source = source with { Content = new([scripts]) };
        var deeper = source with { Content = new([scripts with
            { Base = new([fraction with { Numerator = new([fraction]) }]) }]) };
        TestAssert.True(!DocxMathLayout.TryGetRuns(deeper, CancellationToken.None, out _) &&
            !DocxMathLayout.TryMeasure(deeper, _ => resource, CancellationToken.None, out var box, out _) && box is null,
            "The bounded first nesting grammar cannot admit deeper structures through recursive composition.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => resource, cancelled.Token, out _, out _));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontNoBarStacksKeepSeparatedSourceRolesWithoutPaintingARule()
    {
        byte[] constants = Constants(thickness: -1);
        Set(OpenTypeMathConstant.StackTopDisplayStyleShiftUp, 600);
        Set(OpenTypeMathConstant.StackBottomDisplayStyleShiftDown, 400);
        Set(OpenTypeMathConstant.StackDisplayStyleGapMin, 1800);
        var resource = Resource(TestFontBuilder.CreateMathFont(constants));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        source = source with { Content = new([fraction with { TypeValue = "noBar" }]) };
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
            reason ?? "Stack metrics do not depend on an unused fraction-rule thickness.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal(0, measured.Rules.Count);
        TestAssert.Equal("A+BC", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        double bottomOfTop = double.PositiveInfinity, topOfBottom = double.NegativeInfinity;
        foreach (DocxMathGlyphPlacement glyph in measured.Glyphs)
        {
            TestAssert.True(resource.Font.Font.TryGetGlyphBounds(glyph.OriginalGlyph, out var bounds), "Expected source ink.");
            double scale = glyph.FontSize / resource.Font.Font.UnitsPerEm;
            if (glyph.SourceCodePoint == 'C') topOfBottom = glyph.BaselineY + bounds.YMax * scale;
            else bottomOfTop = Math.Min(bottomOfTop, glyph.BaselineY + bounds.YMin * scale);
        }
        TestAssert.True(bottomOfTop - topOfBottom >= 1800d * 12d / resource.Font.Font.UnitsPerEm - 1e-9,
            "The independently stored ink bounds must respect the font's minimum stack gap.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        TestAssert.True(!graphics.ToString().Contains(" re f", StringComparison.Ordinal), "A stack must not acquire a visible fraction bar.");
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Owned source encodings/resources must form a valid PDF.");
        void Set(OpenTypeMathConstant field, short value) => BinaryPrimitives.WriteInt16BigEndian(constants.AsSpan(18 + (int)field * 4), value);
    }

    public static void MathFontNoBarStacksRejectInvalidMetricsAndRetainInlineScriptOwnership()
    {
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        source = source with { Display = false, Content = new([fraction with { TypeValue = "noBar" }]) };
        var valid = Resource(TestFontBuilder.CreateMathFont(Constants()));
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => valid, CancellationToken.None, out var box, out _), "Expected inline stack.");
        TestAssert.True(TestAssert.NotNull(box).Glyphs.All(g => Math.Abs(g.FontSize - 8.52d) < 1e-9),
            "Stack children retain the existing parent-owned inline script scale.");
        byte[] invalid = Constants();
        BinaryPrimitives.WriteInt16BigEndian(invalid.AsSpan(18 + (int)OpenTypeMathConstant.StackGapMin * 4), -1);
        var bad = Resource(TestFontBuilder.CreateMathFont(invalid));
        TestAssert.True(!DocxMathLayout.TryMeasure(source, _ => bad, CancellationToken.None, out box, out _)
            && box is null, "Invalid stack metrics cannot publish a partial box.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => valid, cancelled.Token, out _, out _));
    }

    public static void MathFontNormalFractionChildrenRetainParentConstantsAndLeadingWithoutChildMathTables()
    {
        OpenTypeFont parent = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(Constants()));
        OpenTypeFont text = TestFontBuilder.LoadTestFont();
        var child = new PdfFontResource("Text1", PdfEmbeddedFont.Create(text, "A+BC".Select(c => (int)c), CancellationToken.None));
        DocxMathRow Normal(string value) => new([new DocxMathRun(new(value, 12d, null, false, false, false, null, "TestFont"),
            null, null) { NormalTextValue = true }]);
        var equation = new DocxEquation(new([new DocxMathFraction(Normal("A+B"), Normal("C"), null)]), 0, 0, false, null, null, null)
            { ParentFontSize = 12d };
        TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => child, CancellationToken.None, out _, out _),
            "Text faces cannot invent parent MATH fraction constants.");
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => child, CancellationToken.None, out var box, out var reason, parent),
            reason ?? "A separately owned parent supplies script scale and fraction constants.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.True(measured.Glyphs.All(glyph => ReferenceEquals(glyph.Font, child) && Math.Abs(glyph.FontSize - 8.52d) < 1e-9),
            "Scaled normal text retains its text-face subset and ordinary glyphs.");
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(measured, CancellationToken.None, out double above, out _),
            "Parent-owned leading survives composition even when no child has MATH.");
        TestAssert.True(Math.Abs(above - measured.AboveBaseline - 120d * 12d / parent.UnitsPerEm) < 1e-9,
            "The parent font contributes one mathematical leading term, independently of child font tables.");
        TestAssert.True(measured.ParentMetrics is { } owner && ReferenceEquals(owner.Font, parent) && owner.Size == 12d,
            "Font and size remain one typed metric owner.");
    }

    public static void MathFontInlineFractionRetainsParentMetricsAndScaledChildGlyphs()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        TestAssert.True(DocxMathLayout.TryMeasure(Fraction() with { Display = false }, _ => resource, CancellationToken.None,
            out DocxMathBox? box, out string? reason), reason ?? "Expected an inline fraction.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.True(measured.Glyphs.All(glyph => Math.Abs(glyph.FontSize - 8.52d) < 1e-9),
            "Numerator and denominator glyphs own the script-style font size.");
        TestAssert.True(Math.Abs(measured.Rules[0].Thickness - 40d * 12d / resource.Font.Font.UnitsPerEm) < 1e-9,
            "The fraction axis/rule metrics remain owned by the parent math size, independently of child scaling.");
        TestAssert.True(measured.Glyphs.Take(3).All(glyph => glyph.BaselineY > measured.Rules[0].Y + measured.Rules[0].Thickness) &&
            measured.Glyphs[3].BaselineY < measured.Rules[0].Y, "Scaled child roles remain on opposite sides of the parent rule.");
    }

    public static void MathFontFractionConstantsDriveStoredRuleAndChildBaselines()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation equation = Fraction();
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out string? reason),
            reason ?? "Expected a display fraction.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal(4, measured.Glyphs.Count);
        TestAssert.Equal(1, measured.Rules.Count);
        DocxMathRulePlacement rule = measured.Rules[0];
        TestAssert.True(measured.Glyphs.Take(3).All(glyph => glyph.BaselineY > rule.Y + rule.Thickness),
            "The numerator must retain its own baseline above the fraction rule.");
        TestAssert.True(measured.Glyphs[3].BaselineY < rule.Y,
            "The denominator must retain its own baseline below the fraction rule.");
        TestAssert.True(Math.Abs(rule.Thickness - 40d * 12d / resource.Font.Font.UnitsPerEm) < 1e-9,
            "Rule thickness must come from the bounded font MATH constants.");
        TestAssert.True(Math.Abs(rule.Width - measured.Width) < 1e-9,
            "The stored rule must cover the measured child-box width.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        TestAssert.Contains(" re f", graphics.ToString());
        TestAssert.True(graphics.UsedFontResourceNames.SetEquals(new[] { "Math1" }), "Fraction painting must use the measured font resource.");
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Stored fraction geometry must produce a validated PDF.");
    }

    public static void MathFontFractionMissingInvalidOrMixedMetricsKeepWholeBoxFallback()
    {
        DocxEquation equation = Fraction();
        foreach (var font in new[] { TestFontBuilder.LoadTestFont(), OpenTypeFont.Load(TestFontBuilder.CreateMathFont(Constants(-1))) })
        {
            var resource = new PdfFontResource("Math1", PdfEmbeddedFont.Create(font, "A+BC".Select(c => (int)c), CancellationToken.None));
            TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out _),
                "Absent or invalid MATH metrics cannot publish an approximate fraction box.");
            TestAssert.True(box is null, "A failed fraction cannot retain a partially measured box.");
        }
        var valid = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxMathFraction fraction = (DocxMathFraction)equation.Content.Children[0];
        DocxMathRun denominator = (DocxMathRun)fraction.Denominator.Children[0];
        DocxEquation mixed = equation with { Content = new([fraction with
        { Denominator = new([denominator with { Text = denominator.Text with { FontSize = 10d } }]) }]) };
        TestAssert.True(!DocxMathLayout.TryMeasure(mixed, _ => valid, CancellationToken.None, out _, out _),
            "Unowned mixed-size fraction metrics require an explicit fallback.");
    }

    public static void MathFontFractionTallInkConsumesFontLeadingOnce()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        TestAssert.True(DocxMathLayout.TryMeasure(Fraction(), _ => resource, CancellationToken.None, out DocxMathBox? box, out _),
            "Expected a measured fraction.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(measured, CancellationToken.None, out double above, out double height),
            "Composed fraction line metrics must be finite.");
        double leading = 120d * 12d / resource.Font.Font.UnitsPerEm;
        TestAssert.True(Math.Abs(above - measured.AboveBaseline - leading) < 1e-9,
            "Tall composed ink consumes one font-leading term, shared by placement and keep measurement.");
        TestAssert.True(height >= above + measured.BelowBaseline, "The line extent must also contain the denominator.");
    }

    public static void MathFontFractionSourceRolesStylesAndCancellationRemainExplicit()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation equation = Fraction();
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out IReadOnlyList<DocxMathRunSource> runs),
            "Display fractions must discover their source runs without flattening numerator/denominator roles.");
        TestAssert.Equal("A+BC", string.Concat(runs.Select(source => source.Run.Text.Text)));
        TestAssert.True(DocxMathLayout.TryGetRuns(equation with { Display = false }, CancellationToken.None, out var inline) &&
            inline.All(source => source.ScriptLevel == 1), "Inline fraction arguments retain their script-style ownership.");
        DocxMathFraction fraction = (DocxMathFraction)equation.Content.Children[0];
        foreach (var unsupported in new[] { fraction with { TypeValue = "skw" }, fraction with { HasUnsupportedProperties = true },
            fraction with { Numerator = new([fraction]) } })
            TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = new([unsupported]) }, CancellationToken.None, out _),
                "Unsupported fraction type, formatting or nested structure remains explicit.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, cancelled.Token, out _, out _));
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out _), "Expected valid fraction.");
        TestAssert.Throws<OperationCanceledException>(() => DocxMathBoxPainter.Paint(TestAssert.NotNull(box), new PdfGraphicsBuilder(), 0d, 0d, cancelled.Token));
    }

    public static void MathFontFractionUsesCumulativeWorkAndRejectsInvalidSourceControls()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 4 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Fraction(), _ => resource, CancellationToken.None, out _, out _));
        foreach ((string properties, bool accepted) in new[]
        {
            ("<m:fPr><m:type m:val='bar'/></m:fPr>", true),
            ("<m:fPr><m:ctrlPr/></m:fPr>", false),
            ("<m:fPr><m:type m:val='skw'/></m:fPr>", false)
        })
        {
            string xml = "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:f>" + properties +
                "<m:num><m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>A</m:t></m:r></m:num>" +
                "<m:den><m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>C</m:t></m:r></m:den></m:f></m:oMath>";
            DocxMathRow row = DocxMathParser.Read(System.Xml.Linq.XElement.Parse(xml),
                (_, text) => new(text, 12d, null, false, false, false, null, "TestFont"), CancellationToken.None);
            TestAssert.Equal(accepted, DocxMathLayout.TryGetRuns(Fraction() with { Content = row }, CancellationToken.None, out _));
            TestAssert.True(row.Children[0] is DocxMathFraction, "Fallback must retain the typed fraction source.");
        }
    }

    private static DocxEquation Fraction() => new(new([new DocxMathFraction(Row("A+B"), Row("C"), null)]), 0, 0, true, "center", null, null);
    private static DocxMathRow Row(string text) => new([new DocxMathRun(new(text, 12d, null, false, false, false, null, "TestFont"), "p", null)]);
    private static PdfFontResource Resource(byte[] bytes) => new("Math1", PdfEmbeddedFont.Create(OpenTypeFont.Load(bytes), "A+BC".Select(c => (int)c), CancellationToken.None));
    private static byte[] Constants(short thickness = 40)
    {
        byte[] bytes = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, 1);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(10), 73);
        Set(OpenTypeMathConstant.AxisHeight, 250);
        Set(OpenTypeMathConstant.MathLeading, 120);
        Set(OpenTypeMathConstant.FractionRuleThickness, thickness);
        Set(OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp, 700);
        Set(OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown, 500);
        Set(OpenTypeMathConstant.FractionNumDisplayStyleGapMin, 100);
        Set(OpenTypeMathConstant.FractionDenomDisplayStyleGapMin, 100);
        Set(OpenTypeMathConstant.FractionNumeratorShiftUp, 500);
        Set(OpenTypeMathConstant.FractionDenominatorShiftDown, 400);
        Set(OpenTypeMathConstant.FractionNumeratorGapMin, 100);
        Set(OpenTypeMathConstant.FractionDenominatorGapMin, 100);
        return bytes;
        void Set(OpenTypeMathConstant field, short value) => BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(18 + (int)field * 4), value);
    }
}
