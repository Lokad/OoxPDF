using System.Runtime.InteropServices;

namespace Lokad.OoxPdf.PdfiumRasterizer;

internal static class PdfiumNative
{
    [DllImport("pdfium")]
    public static extern void FPDF_InitLibrary();

    [DllImport("pdfium")]
    public static extern void FPDF_DestroyLibrary();

    [DllImport("pdfium", CharSet = CharSet.Ansi)]
    public static extern IntPtr FPDF_LoadDocument(string filePath, string? password);

    [DllImport("pdfium")]
    public static extern void FPDF_CloseDocument(IntPtr document);

    [DllImport("pdfium")]
    public static extern int FPDF_GetPageCount(IntPtr document);

    [DllImport("pdfium")]
    public static extern IntPtr FPDF_LoadPage(IntPtr document, int pageIndex);

    [DllImport("pdfium")]
    public static extern void FPDF_ClosePage(IntPtr page);

    [DllImport("pdfium")]
    public static extern float FPDF_GetPageWidthF(IntPtr page);

    [DllImport("pdfium")]
    public static extern float FPDF_GetPageHeightF(IntPtr page);

    [DllImport("pdfium")]
    public static extern IntPtr FPDFBitmap_CreateEx(int width, int height, int format, IntPtr firstScan, int stride);

    [DllImport("pdfium")]
    public static extern void FPDFBitmap_Destroy(IntPtr bitmap);

    [DllImport("pdfium")]
    public static extern void FPDFBitmap_FillRect(IntPtr bitmap, int left, int top, int width, int height, uint color);

    [DllImport("pdfium")]
    public static extern void FPDF_RenderPageBitmap(IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);

    [DllImport("pdfium")]
    public static extern uint FPDF_GetLastError();
}
