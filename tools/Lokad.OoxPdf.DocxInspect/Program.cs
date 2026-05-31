using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.DocxInspect <input.docx> <output-directory>");
    Environment.Exit(2);
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
DocxDocument document = new DocxReader().Read(package);
var renderer = new DocxRenderer();

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};

DocxLayoutSnapshot layout = renderer.InspectLayout(document);
DocxFontPlanSnapshot fontPlan = renderer.InspectFontPlan(document);
File.WriteAllText(
    Path.Combine(outputDirectory, "layout-snapshot.json"),
    JsonSerializer.Serialize(layout, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "font-plan-snapshot.json"),
    JsonSerializer.Serialize(fontPlan, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "page-summary.json"),
    JsonSerializer.Serialize(layout.Pages.Select((page, index) => new
    {
        Page = index + 1,
        page.Width,
        page.Height,
        page.ItemCount,
        page.TextLineCount,
        page.InlineImageCount,
        page.TableRowCount
    }), options));
