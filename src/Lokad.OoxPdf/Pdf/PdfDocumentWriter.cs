using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfDocumentWriter
{
    public static void WriteBlank(Stream stream, IReadOnlyList<PdfPage> pages)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (pages.Count == 0)
        {
            throw new ArgumentException("A PDF document must contain at least one page.", nameof(pages));
        }

        var writer = new PdfObjectWriter(stream);
        writer.WriteHeader();

        int objectCount = 2 + pages.Count * 2;
        writer.WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n");
        writer.WriteObject(2, BuildPagesObject(pages));

        for (int i = 0; i < pages.Count; i++)
        {
            int pageObjectNumber = 3 + i * 2;
            int contentObjectNumber = pageObjectNumber + 1;
            PdfPage page = pages[i];

            writer.WriteObject(pageObjectNumber, FormattableString.Invariant(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {FormatNumber(page.Width)} {FormatNumber(page.Height)}] /Contents {contentObjectNumber} 0 R /Resources << >> >>\n"));
            byte[] contentBytes = Encoding.ASCII.GetBytes(page.Content);
            writer.WriteObject(contentObjectNumber, FormattableString.Invariant(
                $"<< /Length {contentBytes.Length} >>\nstream\n{page.Content}endstream\n"));
        }

        long xrefOffset = writer.Position;
        writer.WriteAscii(FormattableString.Invariant($"xref\n0 {objectCount + 1}\n"));
        writer.WriteAscii("0000000000 65535 f \n");
        foreach (long offset in writer.Offsets)
        {
            writer.WriteAscii(FormattableString.Invariant($"{offset:0000000000} 00000 n \n"));
        }

        writer.WriteAscii(FormattableString.Invariant(
            $"trailer\n<< /Size {objectCount + 1} /Root 1 0 R >>\nstartxref\n{xrefOffset}\n%%EOF\n"));
    }

    private static string BuildPagesObject(IReadOnlyList<PdfPage> pages)
    {
        var builder = new StringBuilder();
        builder.Append("<< /Type /Pages /Count ");
        builder.Append(CultureInfo.InvariantCulture, $"{pages.Count}");
        builder.Append(" /Kids [");
        for (int i = 0; i < pages.Count; i++)
        {
            if (i > 0)
            {
                builder.Append(' ');
            }

            builder.Append(CultureInfo.InvariantCulture, $"{3 + i * 2} 0 R");
        }

        builder.Append("] >>\n");
        return builder.ToString();
    }

    private static string FormatNumber(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
}
