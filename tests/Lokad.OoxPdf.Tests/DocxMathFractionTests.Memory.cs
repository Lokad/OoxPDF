using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathFractionTests
{
    public static void DocxMathAssemblyPreparationRetainsAdmittedCapacityThroughLayout()
    {
        long compact = ObserveAssemblyPreparation(3, 1);
        long large = ObserveAssemblyPreparation(ushort.MaxValue, 1);
        long compound = ObserveAssemblyPreparation(ushort.MaxValue, 2);
        long placementPayloadFloor = (long)ushort.MaxValue * 8;
        TestAssert.True(large - compact >= placementPayloadFloor,
            "Equal-sized font programs must admit their significant font-driven placement expansion before page layout.");
        TestAssert.True(compound - large >= placementPayloadFloor,
            "Two supported structures in one equation retain both assemblies under the same preparation owner.");
    }

    private static long ObserveAssemblyPreparation(int count, int structures)
    {
        var budget = new TestMemoryBudget();
        long retainedBeforeLayout = -1;
        int assemblyAdmissions = 0;
        budget.RefuseChange = (_, bytes) =>
        {
            if (InMathPreparation() && bytes >= (long)count * 8) assemblyAdmissions++;
            if (!InFirstLayout()) return false;
            retainedBeforeLayout = budget.Reserved;
            return true;
        };
        using var input = AssemblyMemoryPackage(structures);
        using var output = new MemoryStream();
        var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = new AssemblyMemoryResolver(AssemblyMemoryFont(count)), MemoryBudget = budget }));
        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, error.Kind);
        TestAssert.True(retainedBeforeLayout > 0 && assemblyAdmissions >= structures,
            "The real public conversion reaches funded assembly construction and retains its capacity at first layout.");
        TestAssert.Equal(0L, output.Length);
        AssertAssemblyMemoryRefunded(budget);
        return retainedBeforeLayout;
    }

    public static void DocxMathAssemblyPreparationRefusalAndCancellationPreservePublication()
    {
        byte[] font = AssemblyMemoryFont(ushort.MaxValue);
        using var package = AssemblyMemoryPackage(1);
        string source = Path.Combine(Path.GetTempPath(), "math-admission-" + Guid.NewGuid().ToString("N") + ".docx");
        File.WriteAllBytes(source, package.ToArray());
        try
        {
            foreach (bool cancel in new[] { false, true })
            foreach (bool file in new[] { false, true })
            {
                using var cancellation = new CancellationTokenSource();
                var budget = new TestMemoryBudget();
                bool reached = false;
                budget.RefuseChange = (_, bytes) =>
                {
                    if (bytes < (long)ushort.MaxValue * 8 || !InMathPreparation()) return false;
                    reached = true;
                    if (!cancel) return true;
                    cancellation.Cancel();
                    return false;
                };
                string target = Path.ChangeExtension(source, ".pdf");
                File.WriteAllBytes(target, [7, 8, 9]);
                using var input = new MemoryStream(package.ToArray());
                using var output = new MemoryStream();
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = new AssemblyMemoryResolver(font), MemoryBudget = budget };
                try
                {
                    void Convert()
                    {
                        if (file) OoxPdfConverter.Convert(source, target, options, cancellation.Token);
                        else OoxPdfConverter.Convert(input, output, options, cancellation.Token);
                    }
                    if (cancel) TestAssert.Throws<OperationCanceledException>(Convert);
                    else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission,
                        TestAssert.Throws<OoxPdfLimitExceededException>(Convert).Kind);
                    TestAssert.True(reached, "Refusal/cancellation occurs in assembly preparation, before allocation and page painting.");
                    TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 7, 8, 9 }), "Failure preserves the existing destination.");
                    TestAssert.Equal(0L, output.Length);
                    AssertAssemblyMemoryRefunded(budget);
                }
                finally { File.Delete(target); }
            }
        }
        finally { File.Delete(source); }
    }

    // Caller-side observation of real producer boundaries; no library hooks.
    private static bool InMathPreparation() => new StackTrace().GetFrames().Any(frame =>
        frame.GetMethod()?.DeclaringType?.Name == "DocxMathGlyphAssembly");
    private static bool InFirstLayout() => new StackTrace().GetFrames().Any(frame =>
        frame.GetMethod()?.Name == "AdmitDocxLayoutPages");

    private static void AssertAssemblyMemoryRefunded(TestMemoryBudget budget)
    {
        TestAssert.Equal(0L, budget.Reserved);
        TestAssert.Equal(budget.ScopesCreated, budget.ScopesDisposed);
        TestAssert.Equal(0, budget.CallerDisposals);
    }

    private static byte[] AssemblyMemoryFont(int count)
    {
        // Same source size, constants and glyphs; only valid assembly count and
        // advance differ. These abbreviated SFNTs test our parser, not conformance.
        byte[] constants = Constants();
        int variants = constants.Length;
        byte[] table = new byte[variants + 40 + ushort.MaxValue * 10];
        constants.CopyTo(table, 0);
        void Set(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(at), value);
        Set(8, (ushort)variants);
        Set(variants + 2, 12); Set(variants + 6, 1); Set(variants + 10, 18);
        Set(variants + 12, 1); Set(variants + 14, 1); Set(variants + 16, 34);
        Set(variants + 18, 16); // assembly relative to construction; zero variants
        Set(variants + 38, (ushort)count);
        for (int index = 0; index < ushort.MaxValue; index++)
        {
            Set(variants + 40 + index * 10, 34);
            Set(variants + 46 + index * 10, (ushort)(count == 3 ? 10000 : 1));
        }
        return TestFontBuilder.CreateMathFont(table);
    }

    private static MemoryStream AssemblyMemoryPackage(int structures)
    {
        var result = new MemoryStream();
        using (var zip = new ZipArchive(result, ZipArchiveMode.Create, true))
        {
            Add("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>");
            Add("_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='main' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");
            const string literal = "<m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>A</m:t></m:r>";
            const string structure = "<m:d><m:dPr><m:begChr m:val='A'/><m:endChr m:val=''/></m:dPr><m:e><m:f><m:num>" + literal + "</m:num><m:den>" + literal + "</m:den></m:f></m:e></m:d>";
            Add("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><w:body><w:p><m:oMath>" + string.Concat(Enumerable.Repeat(structure, structures)) + "</m:oMath></w:p><w:sectPr/></w:body></w:document>");
            void Add(string name, string text) { using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(text); }
        }
        result.Position = 0;
        return result;
    }

    private sealed class AssemblyMemoryResolver(byte[] bytes) : IFontResolver
    {
        private readonly MemoryFontProgramSource source = new("math-assembly-memory-test", bytes);
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "Test Sans",
            new FontStyleKey(HasMathTable: true), source, IsFallback: false);
    }
}
