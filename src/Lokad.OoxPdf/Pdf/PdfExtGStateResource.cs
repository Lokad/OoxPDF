namespace Lokad.OoxPdf.Pdf;

internal readonly record struct PdfExtGStateResource(string ResourceName, double FillAlpha, double StrokeAlpha, PdfLuminositySoftMask? SoftMask = null);
