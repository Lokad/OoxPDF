using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

// Single home for OOXML namespace URIs. Per-file XNamespace copies had
// drifted into near-identical declarations (Drawing x8, Presentation x6),
// so new code should use these instead of redeclaring the strings.
// Names keep the historical Namespace suffix so existing usages compile
// unchanged once a file adds using static OoxNamespaces.
internal static class OoxNamespaces
{
    public static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    public static readonly XNamespace DiagramNamespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    public static readonly XNamespace WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static readonly XNamespace WordprocessingDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public static readonly XNamespace MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    public static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    public static readonly XNamespace VmlNamespace = "urn:schemas-microsoft-com:vml";
    public static readonly XNamespace Office2010WordNamespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    public static readonly XNamespace WordprocessingShapeNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    public static readonly XNamespace Office2012WordNamespace = "http://schemas.microsoft.com/office/word/2012/wordml";
    public static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    public static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
}
