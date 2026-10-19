namespace Lokad.OoxPdf.Docx;

// Printed page coordinates retain the paper center while scaling its extent.
internal readonly record struct DocxPrintedPageFrame(double Height, double Scale)
{
    public double MapY(double sourceY) => sourceY * Scale + (Height / 2d) * (1d - Scale);
    public double UnmapY(double physicalY) => (physicalY - (Height / 2d) * (1d - Scale)) / Scale;

    public double BodyYOffset(double firstBaselineY, double sourceBaselineOffset) =>
        (firstBaselineY - Height / 2d) * (1d - Scale) + sourceBaselineOffset * Scale;
}
