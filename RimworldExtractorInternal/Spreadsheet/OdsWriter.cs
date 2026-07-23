using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace RimworldExtractorInternal.Spreadsheet;

public static class OdsWriter
{
    private static readonly XNamespace OfficeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace TableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace TextNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace ManifestNs = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

    public static void SaveGrid(string filePath, Grid grid)
    {
        if (File.Exists(filePath)) File.Delete(filePath);

        using var zip = ZipFile.Open(filePath, ZipArchiveMode.Create);

        // 1. mimetype 생성
        var mimetypeEntry = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
        using (var writer = new StreamWriter(mimetypeEntry.Open(), Encoding.ASCII))
        {
            writer.Write("application/vnd.oasis.opendocument.spreadsheet");
        }

        // 2. META-INF/manifest.xml 생성
        var manifestEntry = zip.CreateEntry("META-INF/manifest.xml");
        using (var stream = manifestEntry.Open())
        {
            var manifestDoc = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement(ManifestNs + "manifest",
                    new XAttribute(XNamespace.Xmlns + "manifest", ManifestNs.NamespaceName),
                    new XAttribute(ManifestNs + "version", "1.2"),
                    new XElement(ManifestNs + "file-entry",
                        new XAttribute(ManifestNs + "full-path", "/"),
                        new XAttribute(ManifestNs + "version", "1.2"),
                        new XAttribute(ManifestNs + "media-type", "application/vnd.oasis.opendocument.spreadsheet")),
                    new XElement(ManifestNs + "file-entry",
                        new XAttribute(ManifestNs + "full-path", "content.xml"),
                        new XAttribute(ManifestNs + "media-type", "text/xml"))
                )
            );
            manifestDoc.Save(stream);
        }

        // 3. content.xml 생성 (Grid 데이터 기록)
        var contentEntry = zip.CreateEntry("content.xml");
        using (var stream = contentEntry.Open())
        {
            var tableElement = new XElement(TableNs + "table", 
                new XAttribute(TableNs + "name", string.IsNullOrEmpty(grid.Name) ? "Sheet1" : grid.Name));

            foreach (var rowData in grid.Rows)
            {
                var rowElement = new XElement(TableNs + "table-row");
                foreach (var cellValue in rowData)
                {
                    rowElement.Add(new XElement(TableNs + "table-cell",
                        new XAttribute(OfficeNs + "value-type", "string"),
                        new XElement(TextNs + "p", cellValue ?? string.Empty)));
                }
                tableElement.Add(rowElement);
            }

            var contentDoc = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement(OfficeNs + "document-content",
                    new XAttribute(XNamespace.Xmlns + "office", OfficeNs.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "table", TableNs.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "text", TextNs.NamespaceName),
                    new XAttribute(OfficeNs + "version", "1.2"),
                    new XElement(OfficeNs + "body",
                        new XElement(OfficeNs + "spreadsheet", tableElement))
                )
            );
            contentDoc.Save(stream);
        }
    }
}