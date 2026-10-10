using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class PdfTransparencyGroupTests
{
    public static void TransparencyGroupsUseLocalResourcesAndSingleIsolation()
    {
        var first = new PdfGraphicsBuilder();
        first.SetAlpha(.25, .5);
        first.SetFillRgb(255, 0, 0);
        first.FillRectangle(10, 10, 50, 50);
        var second = new PdfGraphicsBuilder();
        second.SetAlpha(.75, 1);
        second.SetFillRgb(0, 0, 255);
        second.FillRectangle(20, 20, 50, 50);
        var parent = new PdfGraphicsBuilder();
        parent.DrawTransparencyGroup(first, Bounds, .5);
        parent.DrawTransparencyGroup(second, Bounds, .25);
        string pdf = Encoding.ASCII.GetString(Write([Page(parent)]));
        TestAssert.Equal(2, Count(pdf, "/Subtype /Form"));
        TestAssert.Equal(2, Count(pdf, "/I true /K false"));
        TestAssert.Contains("/ca 0.25 /CA 0.5", pdf);
        TestAssert.Contains("/ca 0.75 /CA 1", pdf);
        TestAssert.Contains("/ca 0.5 /CA 0.5", pdf);
        TestAssert.Contains("/" + first.ExtGStates[0].ResourceName + " gs", pdf);
        TestAssert.Contains("/" + second.ExtGStates[0].ResourceName + " gs", pdf);
        TestAssert.Contains("/Tr1 Do", pdf);
        TestAssert.Contains("/Tr2 Do", pdf);
    }

    public static void TransparencyGroupsSharePageSpillWindowAndBlankDescriptors()
    {
        PdfPage[] pages = [Page(Nested(20_000)), new PdfPage(100, 100, "0 g\n")];
        byte[] resident = Write(pages);
        using PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages(pages,
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, CancellationToken.None);
        TestAssert.Equal(2, staged.PageEntries.Count);
        TestAssert.Equal(2, staged.GroupEntries.Count);
        TestAssert.Equal(4, staged.Staging.Count);
        TestAssert.True(staged.Staging.SpilledBytes > 0, "Group and page streams must enter the same spill store.");
        TestAssert.True(staged.Pages.All(page => page.Content.Length == 0) && staged.GroupEntries.Keys.All(group => group.Content.Length == 0),
            "Completed payload strings must leave staged descriptors.");
        using var output = new MemoryStream();
        PdfDocumentWriter.EmitStaged(output, staged, CancellationToken.None);
        TestAssert.True(resident.AsSpan().SequenceEqual(output.ToArray()), "Spill output must retain all stream bytes and references.");
    }

    public static void TransparencyGroupsChargeNestedContentOnceBeforeAdmission()
    {
        var leaf = new PdfGraphicsBuilder();
        leaf.SetFillRgb(0, 0, 255);
        leaf.FillRectangle(0, 0, 100, 100);
        long leafLength = leaf.ToString().Length;
        var middle = new PdfGraphicsBuilder();
        middle.DrawTransparencyGroup(leaf, Bounds, .5);
        long middleLength = middle.ToString().Length;
        var root = new PdfGraphicsBuilder();
        root.DrawTransparencyGroup(middle, Bounds, .5);
        long pageLength = root.ToString().Length;
        long exact = leafLength + middleLength + pageLength;
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxPdfContentBytesPerConversion = exact }))
        {
            var scopedMiddle = new PdfGraphicsBuilder();
            scopedMiddle.DrawTransparencyGroup(leaf, Bounds, .5);
            var scopedRoot = new PdfGraphicsBuilder();
            scopedRoot.DrawTransparencyGroup(scopedMiddle, Bounds, .5);
            scope.Budget.ChargePdfContentBytes(scopedRoot.ToString().Length);
            TestAssert.Equal(exact, scope.Budget.PdfContentBytes);
            Write([Page(scopedRoot)]);
            TestAssert.Equal(exact, scope.Budget.PdfContentBytes);
            TestAssert.Throws<OoxPdfLimitExceededException>(() => scope.Budget.ChargePdfContentBytes(1));
        }
        using (OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxPdfContentBytesPerConversion = leafLength - 1 }))
        {
            var rejected = new PdfGraphicsBuilder();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => rejected.DrawTransparencyGroup(leaf, Bounds, .5));
            TestAssert.Equal(0, rejected.Groups.Count);
            TestAssert.Equal(string.Empty, rejected.ToString());
        }
    }

    public static void TransparencyGroupsRetainExactOutputAdmission()
    {
        PdfPage[] pages = [Page(Nested())];
        byte[] expected = Write(pages);
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxOutputBytesPerConversion = expected.Length }))
        {
            TestAssert.True(expected.AsSpan().SequenceEqual(Write(pages)), "Exact output limit must admit the complete document.");
            TestAssert.Equal(expected.Length, scope.Budget.PdfOutputBytes);
        }
        foreach (long limit in new long[] { 0, expected.Length - 1 })
        {
            using OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxOutputBytesPerConversion = limit });
            using var output = new MemoryStream();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteBlank(output, pages, CancellationToken.None));
            TestAssert.True(output.Length <= limit, "Group dictionary/stream bytes must be admitted before writing.");
        }
    }

    public static void TransparencyGroupsRollbackContentResourcesAndNames()
    {
        var parent = Nested();
        string before = parent.ToString();
        PdfGraphicsBuilder.ContentMark mark = parent.MarkContent();
        parent.DrawTransparencyGroup(Nested(), Bounds, .25);
        parent.TruncateContent(mark);
        TestAssert.Equal(before, parent.ToString());
        TestAssert.Equal(1, parent.Groups.Count);
        parent.DrawTransparencyGroup(Nested(), Bounds, .75);
        TestAssert.Equal("Tr2", parent.Groups[1].ResourceName);
        string pdf = Encoding.ASCII.GetString(Write([Page(parent)]));
        TestAssert.DoesNotContain("/ca 0.25 /CA 0.25", pdf);
        TestAssert.Contains("/Tr2 Do", pdf);
    }

    public static void TransparencyGroupsRejectUnknownOrUnbalancedContentBeforeOutput()
    {
        foreach (string content in new[] { "/Missing Do\n", "/Missing gs\n", "/Missing sh\n", "Q\n", "q\n", "BT\n", "é\n" })
        {
            var group = new PdfTransparencyGroup(Bounds, content, [], [], []);
            using var output = new MemoryStream();
            TestAssert.Throws<InvalidDataException>(() => PdfDocumentWriter.WriteBlank(output, [GroupPage(group)], CancellationToken.None));
            TestAssert.Equal(0, output.Length);
        }
    }

    public static void TransparencyGroupsRejectDuplicateResourceNames()
    {
        var leaf = new PdfTransparencyGroup(Bounds, "0 g\n", [], [], []);
        var duplicate = new PdfTransparencyGroup(Bounds, "/Tr1 Do\n", [], [], [new("Tr1", leaf), new("Tr1", leaf)]);
        TestAssert.Throws<InvalidDataException>(() => Write([GroupPage(duplicate)]));
        var duplicatePage = new PdfPage(100, 100, "/Tr1 Do\n", [], [], [], [], [], [], groups: [new("Tr1", leaf), new("Tr1", leaf)]);
        TestAssert.Throws<InvalidDataException>(() => Write([duplicatePage]));
    }

    public static void TransparencyGroupsRejectNonFiniteOrEmptyBounds()
    {
        foreach (PdfRectangle bounds in new[] { new PdfRectangle(double.NaN, 0, 100, 100), new(double.PositiveInfinity, 0, 100, 100), new(0, 0, 0, 100), new(0, 0, 100, -1), new(double.MaxValue, 0, double.MaxValue, 100) })
        {
            TestAssert.Throws<InvalidDataException>(() => Write([GroupPage(new PdfTransparencyGroup(bounds, "0 g\n", [], [], []))]));
        }
    }

    public static void TransparencyGroupsBoundNestingBeforeEmission()
    {
        var group = new PdfTransparencyGroup(Bounds, "0 g\n", [], [], []);
        for (int depth = 1; depth < PdfTransparencyGroup.MaxDepth; depth++)
        {
            group = new PdfTransparencyGroup(Bounds, "/Tr1 Do\n", [], [], [new("Tr1", group)]);
        }
        TestAssert.Equal(32, Count(Encoding.ASCII.GetString(Write([GroupPage(group)])), "/Subtype /Form"));
        var tooDeep = new PdfTransparencyGroup(Bounds, "/Tr1 Do\n", [], [], [new("Tr1", group)]);
        TestAssert.Throws<InvalidDataException>(() => Write([GroupPage(tooDeep)]));
    }

    public static void TransparencyGroupsCancellationDisposesSpilledContent()
    {
        string[] before = Directory.GetFiles(Path.GetTempPath(), "OoxPdfPages-*.tmp");
        using var cancellation = new CancellationTokenSource();
        PdfPage page = Page(Nested());
        IEnumerable<PdfPage> CancelAfterPage()
        {
            yield return page;
            cancellation.Cancel();
            yield return page;
        }
        TestAssert.Throws<OperationCanceledException>(() => PdfDocumentWriter.ProduceStagedPages(CancelAfterPage(),
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, cancellation.Token));
        TestAssert.True(before.Order(StringComparer.Ordinal).SequenceEqual(Directory.GetFiles(Path.GetTempPath(), "OoxPdfPages-*.tmp").Order(StringComparer.Ordinal)), "Cancelled production must close and remove its spill file.");
        using PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages([page],
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, CancellationToken.None);
        using var output = new MemoryStream();
        TestAssert.Throws<OperationCanceledException>(() => PdfDocumentWriter.EmitStaged(output, staged, cancellation.Token));
        TestAssert.Equal(0, output.Length);
        staged.Dispose();
        TestAssert.Throws<ObjectDisposedException>(() => PdfDocumentWriter.EmitStaged(output, staged, CancellationToken.None));
    }

    public static void TransparencyGroupsSnapshotResourcesWithoutSharedMutableOwnership()
    {
        var states = new List<PdfExtGStateResource> { new("Gs1", .5, .5, null) };
        var group = new PdfTransparencyGroup(Bounds, "/Gs1 gs\n0 g\n", states, [], []);
        states.Clear();
        string pdf = Encoding.ASCII.GetString(Write([GroupPage(group)]));
        TestAssert.Contains("/ca 0.5 /CA 0.5", pdf);
        TestAssert.Equal(1, group.ExtGStates.Count);
    }

    public static void TransparencyGroupsSharedReferencesStageAndEmitOnce()
    {
        var leaf = new PdfTransparencyGroup(Bounds, "0 g\n", [], [], []);
        PdfPage[] pages = [GroupPage(leaf), GroupPage(leaf)];
        using PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages(pages, new OoxConversionLimits(), null, CancellationToken.None);
        TestAssert.Equal(1, staged.GroupEntries.Count);
        using var output = new MemoryStream();
        PdfDocumentWriter.EmitStaged(output, staged, CancellationToken.None);
        TestAssert.Equal(1, Count(Encoding.ASCII.GetString(output.ToArray()), "/Subtype /Form"));
    }

    public static void TransparencyGroupsStreamingStagingReleasesEarlierPayloadOwners()
    {
        WeakReference? earlier = null;
        IEnumerable<PdfPage> Pages()
        {
            yield return WeakGroupPage(out earlier);
            yield return new PdfPage(100, 100, "0 g\n");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            TestAssert.True(earlier is { IsAlive: false }, "Staging must release earlier source groups and their content strings.");
            yield return new PdfPage(100, 100, "0 g\n");
        }
        using PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages(Pages(),
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, CancellationToken.None);
        TestAssert.Equal(3, staged.Pages.Count);
        TestAssert.Equal(1, staged.GroupEntries.Count);
        using var output = new MemoryStream();
        PdfDocumentWriter.EmitStaged(output, staged, CancellationToken.None);
        TestAssert.Contains("/Subtype /Form", Encoding.ASCII.GetString(output.ToArray()));
    }

    public static void TransparencyGroupsResolveEqualNamesInIndependentNamespaces()
    {
        var first = new PdfTransparencyGroup(Bounds, "/GS1 gs\n1 0 0 rg\n0 0 50 50 re f\n", [new("GS1", .25, .25, null)], [], []);
        var second = new PdfTransparencyGroup(Bounds, "/GS1 gs\n0 0 1 rg\n25 25 50 50 re f\n", [new("GS1", .75, .75, null)], [], []);
        var page = new PdfPage(100, 100, "/Tr1 Do\n/Tr2 Do\n", [], [], [], [], [], [], groups: [new("Tr1", first), new("Tr2", second)]);
        string pdf = Encoding.ASCII.GetString(Write([page]));
        TestAssert.Equal(2, Count(pdf, "/GS1 gs"));
        TestAssert.Contains("/GS1 << /ca 0.25 /CA 0.25", pdf);
        TestAssert.Contains("/GS1 << /ca 0.75 /CA 0.75", pdf);
    }


    public static void VectorSoftMasksBindLocalFormsAcrossIndependentNamespaces()
    {
        var first = Masked(.25);
        var second = Masked(.75);
        var enclosing = new PdfGraphicsBuilder();
        enclosing.DrawTransparencyGroup(second, Bounds, .5);
        string pdf = Encoding.ASCII.GetString(Write([Page(first), Page(enclosing)]));
        var bindings = System.Text.RegularExpressions.Regex.Matches(pdf,
            @"/XObject << /Tr1 (\d+) 0 R >> /ExtGState << /GSV1 << /ca (?:0.25|0.75) /CA 1 /SMask << /S /Luminosity /G (\d+) 0 R");
        TestAssert.Equal(2, bindings.Count);
        foreach (System.Text.RegularExpressions.Match binding in bindings)
        {
            TestAssert.Equal(binding.Groups[1].Value, binding.Groups[2].Value);
        }
        TestAssert.True(bindings[0].Groups[1].Value != bindings[1].Groups[1].Value,
            "Equal local mask names must resolve independently in page and parent Form resources.");
        TestAssert.Equal(3, Count(pdf, "/Subtype /Form"));
        TestAssert.Equal(2, Count(pdf, "/S /Luminosity"));
        TestAssert.Equal(1, System.Text.RegularExpressions.Regex.Matches(pdf, @" Do\r?\n").Count);
        TestAssert.Contains("/BC [0 0 0]", pdf);
    }

    public static void VectorSoftMasksAdmitOnceAndRollbackResourcesTogether()
    {
        PdfGraphicsBuilder mask = Mask();
        long length = mask.ToString().Length;
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxPdfContentBytesPerConversion = length }))
        {
            var parent = new PdfGraphicsBuilder();
            parent.SetVectorLuminositySoftMask(mask, Bounds, .5, 1);
            TestAssert.Equal(length, scope.Budget.PdfContentBytes);
            TestAssert.Throws<OoxPdfLimitExceededException>(() => parent.SetVectorLuminositySoftMask(mask, Bounds, .5, 1));
            TestAssert.Equal(1, parent.ExtGStates.Count);
            TestAssert.Equal(1, parent.Groups.Count);
        }
        using (OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxPdfContentBytesPerConversion = length - 1 }))
        {
            var rejected = new PdfGraphicsBuilder();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => rejected.SetVectorLuminositySoftMask(mask, Bounds, .5, 1));
            TestAssert.Equal(string.Empty, rejected.ToString());
            TestAssert.Equal(0, rejected.Groups.Count);
            TestAssert.Equal(0, rejected.ExtGStates.Count);
        }
        var builder = new PdfGraphicsBuilder();
        builder.SetFillRgb(0, 0, 255);
        var mark = builder.MarkContent();
        builder.SetVectorLuminositySoftMask(mask, Bounds, .5, 1);
        builder.TruncateContent(mark);
        TestAssert.Equal(0, builder.Groups.Count);
        TestAssert.Equal(0, builder.ExtGStates.Count);
        builder.SetVectorLuminositySoftMask(mask, Bounds, .75, 1);
        TestAssert.Equal("Tr1", builder.ExtGStates.Single().LuminosityGroupName!);
        TestAssert.Equal("GSV1", builder.ExtGStates.Single().ResourceName);
        Write([Page(builder)]);
    }

    public static void VectorSoftMasksShareSpillWindowWithoutRetainingSourceContent()
    {
        PdfPage page = Page(Masked(.5, 30_000));
        byte[] resident = Write([page]);
        using PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages([page],
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, CancellationToken.None);
        TestAssert.Equal(2, staged.Staging.Count);
        TestAssert.Equal(1, staged.GroupEntries.Count);
        TestAssert.True(staged.Staging.SpilledBytes > 65_536, "Large vector masks must spill through the shared content store.");
        TestAssert.Equal(string.Empty, staged.Pages.Single().Content);
        TestAssert.Equal(string.Empty, staged.GroupEntries.Keys.Single().Content);
        TestAssert.Equal("Tr1", staged.Pages.Single().ExtGStates.Single().LuminosityGroupName!);
        using var output = new MemoryStream();
        PdfDocumentWriter.EmitStaged(output, staged, CancellationToken.None);
        TestAssert.True(resident.AsSpan().SequenceEqual(output.ToArray()), "Spilled mask references and content must retain bytes.");
    }

    public static void VectorSoftMasksRejectForeignOrAmbiguousBindingsBeforeOutput()
    {
        var local = new PdfTransparencyGroup(Bounds, "0 g\n", [], [], []);
        var image = PdfImageXObject.RgbPng(1, 1, [0, 0, 0], null);
        var imageMask = new PdfLuminositySoftMask(image, 0, 0, 100, 100, 0, 0, 0, 0);
        foreach (PdfPage page in new[]
        {
            new PdfPage(100, 100, "/GS1 gs\n", [], [], [new("GS1", 1, 1, null, "Foreign")], [], [], [], groups: [new("Tr1", local)]),
            new PdfPage(100, 100, "/GS1 gs\n", [], [], [new("GS1", 1, 1, imageMask, "Tr1")], [], [], [], groups: [new("Tr1", local)]),
            GroupPage(new PdfTransparencyGroup(Bounds, "/GS1 gs\n", [new("GS1", 1, 1, null, "Tr1")], [], []))
        })
        {
            using var output = new MemoryStream();
            TestAssert.Throws<InvalidDataException>(() => PdfDocumentWriter.WriteBlank(output, [page], CancellationToken.None));
            TestAssert.Equal(0, output.Length);
        }
    }

    public static void VectorSoftMasksReleaseEarlierPayloadOwnersDuringStreamingStaging()
    {
        WeakReference? earlier = null;
        IEnumerable<PdfPage> Pages()
        {
            yield return WeakMaskPage(out earlier);
            yield return new PdfPage(100, 100, "0 g\n");
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            TestAssert.True(earlier is { IsAlive: false }, "Mask graphics-state bindings must not retain earlier source Form payloads.");
            yield return new PdfPage(100, 100, "0 g\n");
        }
        using PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages(Pages(),
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, CancellationToken.None);
        TestAssert.Equal(1, staged.GroupEntries.Count);
        using var output = new MemoryStream();
        PdfDocumentWriter.EmitStaged(output, staged, CancellationToken.None);
        TestAssert.Contains("/S /Luminosity", Encoding.ASCII.GetString(output.ToArray()));
    }

    public static void VectorSoftMasksRetainExactOutputAdmissionAndCancellationCleanup()
    {
        PdfPage page = Page(Masked(.5, 30_000));
        byte[] expected = Write([page]);
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxOutputBytesPerConversion = expected.Length }))
        {
            TestAssert.True(expected.AsSpan().SequenceEqual(Write([page])), "An exact output budget must admit every mask byte once.");
            TestAssert.Equal(expected.Length, scope.Budget.PdfOutputBytes);
        }
        foreach (long maximum in new long[] { 0, expected.Length - 1 })
        {
            using OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxOutputBytesPerConversion = maximum });
            using var output = new MemoryStream();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteBlank(output, [page], CancellationToken.None));
            TestAssert.True(output.Length <= maximum, "Mask streams cannot write beyond the admitted output.");
        }
        string[] before = Directory.GetFiles(Path.GetTempPath(), "OoxPdfPages-*.tmp");
        using var cancellation = new CancellationTokenSource();
        IEnumerable<PdfPage> Pages()
        {
            yield return page;
            cancellation.Cancel();
            yield return page;
        }
        TestAssert.Throws<OperationCanceledException>(() => PdfDocumentWriter.ProduceStagedPages(Pages(),
            new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, cancellation.Token));
        TestAssert.True(before.Order(StringComparer.Ordinal).SequenceEqual(Directory.GetFiles(Path.GetTempPath(), "OoxPdfPages-*.tmp").Order(StringComparer.Ordinal)),
            "Cancelled mask staging must remove its owned spill file.");
    }

    private static PdfGraphicsBuilder Mask(int repetitions = 0)
    {
        var mask = new PdfGraphicsBuilder();
        mask.PaintRadialShading([new(0, 0, 0, 0), new(1, 255, 255, 255)]);
        for (int i = 0; i < repetitions; i++) { mask.FillRectangle(0, 0, 100, 100); }
        return mask;
    }
    private static PdfGraphicsBuilder Masked(double alpha, int repetitions = 0)
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SaveState();
        graphics.SetVectorLuminositySoftMask(Mask(repetitions), Bounds, alpha, 1);
        graphics.SetFillRgb(255, 0, 0);
        graphics.FillRectangle(0, 0, 100, 100);
        graphics.RestoreState();
        return graphics;
    }
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static PdfPage WeakMaskPage(out WeakReference reference)
    {
        var mask = new PdfTransparencyGroup(Bounds, string.Concat(Enumerable.Repeat("0 g\n", 40_000)), [], [], []);
        reference = new WeakReference(mask);
        return new PdfPage(100, 100, "/GS1 gs\n0 0 100 100 re f\n", [], [], [new("GS1", 1, 1, null, "Tr1")], [], [], [], groups: [new("Tr1", mask)]);
    }

    private static PdfRectangle Bounds => new(0, 0, 100, 100);

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static PdfPage WeakGroupPage(out WeakReference reference)
    {
        var group = new PdfTransparencyGroup(Bounds, string.Concat(Enumerable.Repeat("0 g\n", 40_000)), [], [], []);
        reference = new WeakReference(group);
        return GroupPage(group);
    }
    private static PdfGraphicsBuilder Nested(int repetitions = 0)
    {
        var leaf = new PdfGraphicsBuilder();
        leaf.SetFillRgb(255, 0, 0);
        leaf.FillRectangle(0, 0, 100, 100);
        for (int i = 0; i < repetitions; i++) { leaf.FillRectangle(0, 0, 100, 100); }
        var middle = new PdfGraphicsBuilder();
        middle.DrawTransparencyGroup(leaf, Bounds, .5);
        var root = new PdfGraphicsBuilder();
        root.DrawTransparencyGroup(middle, Bounds, .5);
        return root;
    }
    private static PdfPage Page(PdfGraphicsBuilder graphics) => new(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings, [], [], groups: graphics.Groups);
    private static PdfPage GroupPage(PdfTransparencyGroup group) => new(100, 100, "/Tr1 Do\n", [], [], [], [], [], [], groups: [new("Tr1", group)]);
    private static byte[] Write(IReadOnlyList<PdfPage> pages)
    {
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, pages, CancellationToken.None);
        return output.ToArray();
    }
    private static int Count(string text, string value) => (text.Length - text.Replace(value, string.Empty, StringComparison.Ordinal).Length) / value.Length;
}
