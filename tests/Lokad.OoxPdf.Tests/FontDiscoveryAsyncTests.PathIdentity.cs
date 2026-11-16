using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontDiscoveryAsyncTests
{
    public static void FontDiscoveryPreservesCaseDistinctFilePrograms()
    {
        InDirectory(directory =>
        {
            byte[] upper = TestFontBuilder.CreateFontWithFamilyNames((0x0409, "UpperFamily"));
            byte[] lower = TestFontBuilder.CreateFontWithFamilyNames((0x0409, "LowerFamily"));
            File.WriteAllBytes(Path.Combine(directory, "Face.ttf"), upper);
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), lower);
            if (Directory.GetFiles(directory).Length != 2)
                TestAssert.Skip("Case-distinct file control requires a case-sensitive directory.");

            foreach (bool owned in new[] { false, true })
            {
                var pool = new TestMemoryBudget();
                using (var resolver = owned
                    ? CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult()
                    : CreateAsync(directory).AsTask().GetAwaiter().GetResult())
                {
                    var upperFace = resolver.Resolve(new("upperfamily"));
                    var lowerFace = resolver.Resolve(new("LOWERFAMILY"));
                    TestAssert.True(!upperFace.IsFallback && !lowerFace.IsFallback, "Both independently discovered families must resolve exactly.");
                    TestAssert.True(upperFace.Source.StableId != lowerFace.Source.StableId, "Case-distinct files retain independent program ownership.");
                    foreach (var (face, bytes) in new[] { (upperFace, upper), (lowerFace, lower), (upperFace, upper), (lowerFace, lower) })
                    {
                        using var lease = face.Source.AcquireAsync(null).AsTask().GetAwaiter().GetResult();
                        TestAssert.True(lease.Bytes.Span.SequenceEqual(bytes), "Cold and cached acquisition must return the requested source bytes.");
                    }
                }
                TestAssert.Equal(0L, pool.Reserved);
            }
        });
    }

    public static void FontDiscoveryPreservesCaseDistinctDirectorySnapshots()
    {
        InDirectory(directory =>
        {
            string upper = Path.Combine(directory, "Catalog"), lower = Path.Combine(directory, "catalog");
            Directory.CreateDirectory(upper);
            Directory.CreateDirectory(lower);
            if (Directory.GetDirectories(directory).Length != 2)
                TestAssert.Skip("Case-distinct discovery roots require a case-sensitive directory.");
            byte[] upperBytes = TestFontBuilder.CreateFontWithFamilyNames((0x0409, "UpperCatalog"));
            byte[] lowerBytes = TestFontBuilder.CreateFontWithFamilyNames((0x0409, "LowerCatalog"));
            File.WriteAllBytes(Path.Combine(upper, "face.ttf"), upperBytes);
            File.WriteAllBytes(Path.Combine(lower, "face.ttf"), lowerBytes);
            using var first = CreateAsync(upper).AsTask().GetAwaiter().GetResult();
            using var second = CreateAsync(lower).AsTask().GetAwaiter().GetResult();
            foreach (var (resolver, family, bytes) in new[] { (first, "UpperCatalog", upperBytes), (second, "LowerCatalog", lowerBytes) })
            {
                var face = resolver.Resolve(new(family));
                TestAssert.True(!face.IsFallback, "A different-case root must not reuse another directory's metadata.");
                using var lease = face.Source.AcquireAsync(null).AsTask().GetAwaiter().GetResult();
                TestAssert.True(lease.Bytes.Span.SequenceEqual(bytes), "Each catalog retains its own program.");
            }
        });
    }

    public static void FontDiscoveryFamilyNamesRemainCaseInsensitive()
    {
        InDirectory(directory =>
        {
            byte[] bytes = TestFontBuilder.CreateFontWithFamilyNames((0x0409, "MixedFamily"));
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), bytes);
            using var resolver = CreateAsync(directory).AsTask().GetAwaiter().GetResult();
            var first = resolver.Resolve(new("mixedfamily"));
            var second = resolver.Resolve(new("MIXEDFAMILY"));
            TestAssert.True(!first.IsFallback && !second.IsFallback, "Family lookup preserves its independent case-insensitive semantics.");
            TestAssert.Equal(first.Source.StableId, second.Source.StableId);
            using var lease = second.Source.AcquireAsync(null).AsTask().GetAwaiter().GetResult();
            TestAssert.True(lease.Bytes.Span.SequenceEqual(bytes), "Family spelling does not change the selected program.");
        });
    }
}
