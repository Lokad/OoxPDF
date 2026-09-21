using System.Globalization;
using System.Runtime.InteropServices;
using Lokad.OoxPdf.PdfiumRasterizer;

if (args.Length is < 2 or > 3)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.PdfiumRasterizer <input.pdf> <output-directory> [dpi]");
    return 2;
}

string inputPdf = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
// PLAN Q06: unbounded DPI turns page size into a giant pinned bitmap. Cap the range
// generously (all tracked manifests use 144) and fail with usage text, not a crash.
const int MinDpi = 36;
const int MaxDpi = 600;
if (!int.TryParse(args.Length == 3 ? args[2] : "144", CultureInfo.InvariantCulture, out int dpi) ||
    dpi < MinDpi || dpi > MaxDpi)
{
    Console.Error.WriteLine($"Invalid DPI '{(args.Length == 3 ? args[2] : "144")}': expected an integer {MinDpi}-{MaxDpi}.");
    return 2;
}

if (!File.Exists(inputPdf))
{
    Console.Error.WriteLine($"Input PDF was not found: {inputPdf}");
    return 1;
}

Directory.CreateDirectory(outputDirectory);

string nativeDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "vendor", "pdfium", "win-x64", "bin"));
NativeLibrary.SetDllImportResolver(typeof(PdfiumNative).Assembly, (_, _, _) =>
{
    string dllPath = Path.Combine(nativeDirectory, "pdfium.dll");
    return NativeLibrary.Load(dllPath);
});

PdfiumNative.FPDF_InitLibrary();
try
{
    IntPtr document = PdfiumNative.FPDF_LoadDocument(inputPdf, null);
    if (document == IntPtr.Zero)
    {
        Console.Error.WriteLine($"PDFium failed to open PDF. Error code: {PdfiumNative.FPDF_GetLastError()}");
        return 1;
    }

    try
    {
        int pageCount = PdfiumNative.FPDF_GetPageCount(document);
        int failedPages = 0;
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            // PLAN Q06: isolate pages for managed failures so one oversized page does
            // not discard the rest; native crashes remain process-fatal (documented).
            try
            {
                RenderPage(document, pageIndex, Path.Combine(outputDirectory, $"page-{pageIndex + 1:000}.png"), dpi);
            }
            catch (Exception ex) when (ex is InvalidDataException or InvalidOperationException or OutOfMemoryException)
            {
                Console.Error.WriteLine($"Page {pageIndex + 1} failed: {ex.Message}");
                failedPages++;
            }
        }

        if (failedPages != 0)
        {
            Console.Error.WriteLine($"{failedPages} of {pageCount} pages failed.");
            return 1;
        }
    }
    finally
    {
        PdfiumNative.FPDF_CloseDocument(document);
    }
}
finally
{
    PdfiumNative.FPDF_DestroyLibrary();
}

// PLAN Q06: tool peaks are recorded separately from converter peaks.
Console.WriteLine($"Tool peak working set: {System.Diagnostics.Process.GetCurrentProcess().PeakWorkingSet64} bytes.");
return 0;

static void RenderPage(IntPtr document, int pageIndex, string outputPath, int dpi)
{
    IntPtr page = PdfiumNative.FPDF_LoadPage(document, pageIndex);
    if (page == IntPtr.Zero)
    {
        throw new InvalidOperationException($"PDFium failed to load page {pageIndex + 1}.");
    }

    try
    {
        // PLAN Q06: checked dimension math plus explicit pixel/byte caps before the
        // pinned bitmap exists. Mirrors the library pixel budget so rasterizer output
        // always fits what VisualDiff accepts.
        const int MaxRasterDimension = 32768;
        const long MaxRasterPixels = 67_108_864L;
        double scale = dpi / 72d;
        double pageWidth = PdfiumNative.FPDF_GetPageWidthF(page);
        double pageHeight = PdfiumNative.FPDF_GetPageHeightF(page);
        if (!double.IsFinite(pageWidth) || !double.IsFinite(pageHeight) || pageWidth <= 0d || pageHeight <= 0d)
        {
            throw new InvalidDataException($"PDFium reported non-finite page dimensions {pageWidth}x{pageHeight}.");
        }

        long longWidth;
        long longHeight;
        try
        {
            longWidth = checked((long)Math.Ceiling(pageWidth * scale));
            longHeight = checked((long)Math.Ceiling(pageHeight * scale));
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("PDF page dimensions exceed the supported raster size.", ex);
        }

        if (longWidth < 1 || longHeight < 1 || longWidth > MaxRasterDimension || longHeight > MaxRasterDimension)
        {
            throw new InvalidDataException($"PDF page raster dimensions {longWidth}x{longHeight} exceed the maximum of {MaxRasterDimension} per side.");
        }

        long pixelCount;
        long byteCount;
        try
        {
            pixelCount = checked(longWidth * longHeight);
            byteCount = checked(pixelCount * 4L);
        }
        catch (OverflowException ex)
        {
            throw new InvalidDataException("PDF page raster size overflows.", ex);
        }

        if (pixelCount > MaxRasterPixels)
        {
            throw new InvalidDataException($"PDF page raster pixel count {pixelCount} exceeds the maximum of {MaxRasterPixels}.");
        }

        int width = Math.Max(1, (int)longWidth);
        int height = Math.Max(1, (int)longHeight);
        int stride = checked(width * 4);
        byte[] bgra = new byte[byteCount];

        GCHandle handle = GCHandle.Alloc(bgra, GCHandleType.Pinned);
        try
        {
            IntPtr bitmap = PdfiumNative.FPDFBitmap_CreateEx(width, height, 4, handle.AddrOfPinnedObject(), stride);
            if (bitmap == IntPtr.Zero)
            {
                throw new InvalidOperationException("PDFium failed to create a bitmap.");
            }

            try
            {
                PdfiumNative.FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xFFFFFFFF);
                PdfiumNative.FPDF_RenderPageBitmap(bitmap, page, 0, 0, width, height, 0, 0);
                PngWriter.WriteBgra(outputPath, width, height, bgra);
            }
            finally
            {
                PdfiumNative.FPDFBitmap_Destroy(bitmap);
            }
        }
        finally
        {
            handle.Free();
        }
    }
    finally
    {
        PdfiumNative.FPDF_ClosePage(page);
    }
}
