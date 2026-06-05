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
int dpi = args.Length == 3 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 144;
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
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            RenderPage(document, pageIndex, Path.Combine(outputDirectory, $"page-{pageIndex + 1:000}.png"), dpi);
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
        double scale = dpi / 72d;
        int width = Math.Max(1, (int)Math.Ceiling(PdfiumNative.FPDF_GetPageWidthF(page) * scale));
        int height = Math.Max(1, (int)Math.Ceiling(PdfiumNative.FPDF_GetPageHeightF(page) * scale));
        int stride = width * 4;
        byte[] bgra = new byte[stride * height];

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
