using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxContentTypes
{

    private readonly Dictionary<string, string> defaults;
    private readonly Dictionary<string, string> overrides;

    private OoxContentTypes(Dictionary<string, string> defaults, Dictionary<string, string> overrides)
    {
        this.defaults = defaults;
        this.overrides = overrides;
    }

    public static OoxContentTypes Parse(Stream stream, CancellationToken cancellationToken)
    {
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (XElement element in document.Root?.Elements() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.Name == OoxNamespaces.ContentTypesNamespace + "Default")
            {
                string extension = RequiredAttribute(element, "Extension").TrimStart('.');
                AddDeclaration(defaults, extension, element);
            }
            else if (element.Name == OoxNamespaces.ContentTypesNamespace + "Override")
            {
                string partName = OoxPath.NormalizePartName(RequiredAttribute(element, "PartName"));
                AddDeclaration(overrides, partName, element);
            }
        }

        return new OoxContentTypes(defaults, overrides);
    }

    private static void AddDeclaration(Dictionary<string, string> declarations, string key, XElement element)
    {
        string contentType = RequiredAttribute(element, "ContentType");
        if (!declarations.TryAdd(key, contentType) && declarations[key] != contentType)
        {
            throw new InvalidDataException("OOXML package declares conflicting content types.");
        }
    }

    // The package reader validates and canonicalizes each name before lookup.
    public string? GetContentType(string normalizedPartName)
    {
        if (overrides.TryGetValue(normalizedPartName, out string? contentType))
        {
            return contentType;
        }

        string extension = Path.GetExtension(normalizedPartName).TrimStart('.');
        return defaults.TryGetValue(extension, out contentType) ? contentType : null;
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required {element.Name.LocalName} attribute '{name}'.");
    }
}
