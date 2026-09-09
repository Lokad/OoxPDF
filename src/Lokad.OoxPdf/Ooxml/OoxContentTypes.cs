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
                string extension = RequiredAttribute(element, "Extension");
                string defaultKey = extension.TrimStart('.');
                if (defaults.TryGetValue(defaultKey, out string? existingDefault) && existingDefault != RequiredAttribute(element, "ContentType"))
                {
                    throw new InvalidDataException("OOXML package declares conflicting content types.");
                }

                defaults[defaultKey] = RequiredAttribute(element, "ContentType");
            }
            else if (element.Name == OoxNamespaces.ContentTypesNamespace + "Override")
            {
                string partName = OoxPath.NormalizePartName(RequiredAttribute(element, "PartName"));
                string overrideType = RequiredAttribute(element, "ContentType");
                if (overrides.TryGetValue(partName, out string? existingOverride) && existingOverride != overrideType)
                {
                    throw new InvalidDataException("OOXML package declares conflicting content types.");
                }

                overrides[partName] = overrideType;
            }
        }

        return new OoxContentTypes(defaults, overrides);
    }

    public string? GetContentType(string partName)
    {
        string normalized = OoxPath.NormalizePartName(partName);
        if (overrides.TryGetValue(normalized, out string? contentType))
        {
            return contentType;
        }

        string extension = Path.GetExtension(normalized).TrimStart('.');
        return defaults.TryGetValue(extension, out contentType) ? contentType : null;
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required {element.Name.LocalName} attribute '{name}'.");
    }
}
