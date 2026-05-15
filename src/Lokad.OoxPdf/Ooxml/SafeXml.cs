using System.Xml;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal static class SafeXml
{
    public static XmlReaderSettings CreateReaderSettings()
    {
        return new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false
        };
    }

    public static XDocument Load(Stream stream)
    {
        using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings());
        return XDocument.Load(reader, LoadOptions.None);
    }
}
