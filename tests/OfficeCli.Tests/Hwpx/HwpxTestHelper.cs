// File: tests/OfficeCli.Tests/Hwpx/HwpxTestHelper.cs
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace OfficeCli.Tests.Hwpx;

internal static class HwpxTestHelper
{
    // HWPX namespace constants (test-local; production uses HwpxNs.Hp/Hs/Hh/Opf)
    private static readonly XNamespace Hp = "http://www.hancom.co.kr/hwpml/2011/paragraph";
    private static readonly XNamespace Hs = "http://www.hancom.co.kr/hwpml/2011/section";
    private static readonly XNamespace Hh = "http://www.hancom.co.kr/hwpml/2011/head";
    private static readonly XNamespace Hc = "http://www.hancom.co.kr/hwpml/2011/core";
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    /// <summary>
    /// Create a minimal valid HWPX file in memory with one section and one paragraph.
    /// Returns the file path (temp file, caller must clean up).
    /// </summary>
    public static string CreateMinimalHwpx(string text, bool includeHeader = true)
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.hwpx");
        using var stream = File.Create(filePath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        // mimetype
        WriteEntry(archive, "mimetype", "application/hwp+zip");

        // META-INF/container.xml
        var containerXml = new XDocument(
            new XElement("container",
                new XAttribute("version", "1.0"),
                new XElement("rootfiles",
                    new XElement("rootfile",
                        new XAttribute("full-path", "Contents/content.hpf"),
                        new XAttribute("media-type", "application/hwpml-package+xml")))));
        WriteXmlEntry(archive, "META-INF/container.xml", containerXml);

        // Contents/content.hpf (OPF manifest)
        var hpfXml = new XDocument(
            new XElement(Opf + "package",
                new XAttribute(XNamespace.Xmlns + "hpf", Opf.NamespaceName),
                new XElement(Opf + "manifest",
                    new XElement(Opf + "item",
                        new XAttribute("id", "header"),
                        new XAttribute("href", "header.xml"),
                        new XAttribute("media-type", "application/xml")),
                    new XElement(Opf + "item",
                        new XAttribute("id", "section0"),
                        new XAttribute("href", "section0.xml"),
                        new XAttribute("media-type", "application/xml+section"))),
                new XElement(Opf + "spine",
                    new XElement(Opf + "itemref",
                        new XAttribute("idref", "section0")))));
        WriteXmlEntry(archive, "Contents/content.hpf", hpfXml);

        // Contents/header.xml
        if (includeHeader)
        {
            var headerXml = new XDocument(
                new XElement(Hh + "head",
                    new XAttribute(XNamespace.Xmlns + "hh", Hh.NamespaceName),
                    new XElement(Hh + "charProperties",
                        new XElement(Hh + "charPr",
                            new XAttribute("id", "0"),
                            new XAttribute("height", "1000"),
                            // fontRef uses numeric font IDs (not font names).
                            // Resolve via: hangul="0" → <hh:fontface lang="HANGUL"><hh:font id="0" face="...">
                            new XElement(Hh + "fontRef",
                                new XAttribute("hangul", "0"),
                                new XAttribute("latin", "0"),
                                new XAttribute("hanja", "0"),
                                new XAttribute("japanese", "0"),
                                new XAttribute("other", "0"),
                                new XAttribute("symbol", "0"),
                                new XAttribute("user", "0")))),
                    new XElement(Hh + "paraProperties",
                        new XElement(Hh + "paraPr",
                            new XAttribute("id", "0"),
                            new XElement(Hh + "align",
                                new XAttribute("horizontal", "LEFT"),
                                new XAttribute("vertical", "BASELINE")))),
                    new XElement(Hh + "styles",
                        new XAttribute("itemCnt", "2"),
                        new XElement(Hh + "style",
                            new XAttribute("id", "0"),
                            new XAttribute("type", "PARA"),
                            new XAttribute("name", "바탕글"),
                            new XAttribute("engName", "Normal"),
                            new XAttribute("paraPrIDRef", "0"),
                            new XAttribute("charPrIDRef", "0")),
                        new XElement(Hh + "style",
                            new XAttribute("id", "1"),
                            new XAttribute("type", "PARA"),
                            new XAttribute("name", "개요 1"),
                            new XAttribute("engName", "Outline 1"),
                            new XAttribute("paraPrIDRef", "0"),
                            new XAttribute("charPrIDRef", "0")))));
            WriteXmlEntry(archive, "Contents/header.xml", headerXml);
        }

        // Contents/section0.xml
        var sectionXml = new XDocument(
            new XElement(Hs + "sec",
                new XAttribute(XNamespace.Xmlns + "hs", Hs.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "hp", Hp.NamespaceName),
                new XElement(Hp + "p",
                    new XAttribute("paraPrIDRef", "0"),
                    new XAttribute("styleIDRef", "0"),
                    new XElement(Hp + "run",
                        new XAttribute("charPrIDRef", "0"),
                        new XElement(Hp + "t", text)))));
        WriteXmlEntry(archive, "Contents/section0.xml", sectionXml);

        return filePath;
    }

    /// <summary>
    /// Create a HWPX file with multiple sections.
    /// </summary>
    public static string CreateMultiSectionHwpx(params string[][] sectionTexts)
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.hwpx");
        using var stream = File.Create(filePath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteEntry(archive, "mimetype", "application/hwp+zip");

        // META-INF/container.xml
        var containerXml = new XDocument(
            new XElement("container",
                new XAttribute("version", "1.0"),
                new XElement("rootfiles",
                    new XElement("rootfile",
                        new XAttribute("full-path", "Contents/content.hpf"),
                        new XAttribute("media-type", "application/hwpml-package+xml")))));
        WriteXmlEntry(archive, "META-INF/container.xml", containerXml);

        // Build manifest items
        var manifestItems = new List<XElement>
        {
            new(Opf + "item",
                new XAttribute("id", "header"),
                new XAttribute("href", "header.xml"),
                new XAttribute("media-type", "application/xml"))
        };
        var spineItems = new List<XElement>();

        for (int s = 0; s < sectionTexts.Length; s++)
        {
            manifestItems.Add(new XElement(Opf + "item",
                new XAttribute("id", $"section{s}"),
                new XAttribute("href", $"section{s}.xml"),
                new XAttribute("media-type", "application/xml+section")));
            spineItems.Add(new XElement(Opf + "itemref",
                new XAttribute("idref", $"section{s}")));
        }

        var hpfXml = new XDocument(
            new XElement(Opf + "package",
                new XAttribute(XNamespace.Xmlns + "hpf", Opf.NamespaceName),
                new XElement(Opf + "manifest", manifestItems),
                new XElement(Opf + "spine", spineItems)));
        WriteXmlEntry(archive, "Contents/content.hpf", hpfXml);

        // Header
        var headerXml = new XDocument(
            new XElement(Hh + "head",
                new XAttribute(XNamespace.Xmlns + "hh", Hh.NamespaceName),
                new XElement(Hh + "charProperties",
                    new XElement(Hh + "charPr",
                        new XAttribute("id", "0"),
                        new XAttribute("height", "1000"),
                        new XElement(Hh + "fontRef",
                            new XAttribute("hangul", "0"),
                            new XAttribute("latin", "0"),
                            new XAttribute("hanja", "0"),
                            new XAttribute("japanese", "0"),
                            new XAttribute("other", "0"),
                            new XAttribute("symbol", "0"),
                            new XAttribute("user", "0")))),
                new XElement(Hh + "paraProperties",
                    new XElement(Hh + "paraPr",
                        new XAttribute("id", "0"),
                        new XElement(Hh + "align",
                            new XAttribute("horizontal", "LEFT"),
                            new XAttribute("vertical", "BASELINE")))),
                new XElement(Hh + "styles",
                    new XAttribute("itemCnt", "1"),
                    new XElement(Hh + "style",
                        new XAttribute("id", "0"),
                        new XAttribute("type", "PARA"),
                        new XAttribute("name", "바탕글"),
                        new XAttribute("engName", "Normal"),
                        new XAttribute("paraPrIDRef", "0"),
                        new XAttribute("charPrIDRef", "0")))));
        WriteXmlEntry(archive, "Contents/header.xml", headerXml);

        // Sections
        for (int s = 0; s < sectionTexts.Length; s++)
        {
            var paras = sectionTexts[s].Select(t =>
                new XElement(Hp + "p",
                    new XAttribute("paraPrIDRef", "0"),
                    new XAttribute("styleIDRef", "0"),
                    new XElement(Hp + "run",
                        new XAttribute("charPrIDRef", "0"),
                        new XElement(Hp + "t", t))));

            var sectionXml = new XDocument(
                new XElement(Hs + "sec",
                    new XAttribute(XNamespace.Xmlns + "hs", Hs.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "hp", Hp.NamespaceName),
                    paras));
            WriteXmlEntry(archive, $"Contents/section{s}.xml", sectionXml);
        }

        return filePath;
    }

    /// <summary>
    /// Create a HWPX file with a table in the first section.
    /// </summary>
    public static string CreateHwpxWithTable(int rows, int cols,
        bool includeCellMargin = true, bool includeCellSz = true, bool includeCellAddr = true)
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.hwpx");
        using var stream = File.Create(filePath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteEntry(archive, "mimetype", "application/hwp+zip");

        var containerXml = new XDocument(
            new XElement("container",
                new XAttribute("version", "1.0"),
                new XElement("rootfiles",
                    new XElement("rootfile",
                        new XAttribute("full-path", "Contents/content.hpf"),
                        new XAttribute("media-type", "application/hwpml-package+xml")))));
        WriteXmlEntry(archive, "META-INF/container.xml", containerXml);

        var hpfXml = new XDocument(
            new XElement(Opf + "package",
                new XAttribute(XNamespace.Xmlns + "hpf", Opf.NamespaceName),
                new XElement(Opf + "manifest",
                    new XElement(Opf + "item",
                        new XAttribute("id", "header"),
                        new XAttribute("href", "header.xml"),
                        new XAttribute("media-type", "application/xml")),
                    new XElement(Opf + "item",
                        new XAttribute("id", "section0"),
                        new XAttribute("href", "section0.xml"),
                        new XAttribute("media-type", "application/xml+section"))),
                new XElement(Opf + "spine",
                    new XElement(Opf + "itemref",
                        new XAttribute("idref", "section0")))));
        WriteXmlEntry(archive, "Contents/content.hpf", hpfXml);

        var headerXml = new XDocument(
            new XElement(Hh + "head",
                new XAttribute(XNamespace.Xmlns + "hh", Hh.NamespaceName),
                new XElement(Hh + "charProperties",
                    new XElement(Hh + "charPr",
                        new XAttribute("id", "0"),
                        new XAttribute("height", "1000"),
                        new XElement(Hh + "fontRef",
                            new XAttribute("hangul", "0"),
                            new XAttribute("latin", "0"),
                            new XAttribute("hanja", "0"),
                            new XAttribute("japanese", "0"),
                            new XAttribute("other", "0"),
                            new XAttribute("symbol", "0"),
                            new XAttribute("user", "0")))),
                new XElement(Hh + "paraProperties",
                    new XElement(Hh + "paraPr",
                        new XAttribute("id", "0"),
                        new XElement(Hh + "align",
                            new XAttribute("horizontal", "LEFT"),
                            new XAttribute("vertical", "BASELINE")))),
                new XElement(Hh + "styles",
                    new XAttribute("itemCnt", "1"),
                    new XElement(Hh + "style",
                        new XAttribute("id", "0"),
                        new XAttribute("type", "PARA"),
                        new XAttribute("name", "바탕글"),
                        new XAttribute("engName", "Normal"),
                        new XAttribute("paraPrIDRef", "0"),
                        new XAttribute("charPrIDRef", "0")))));
        WriteXmlEntry(archive, "Contents/header.xml", headerXml);

        // Build table
        var trElements = new List<XElement>();
        for (int r = 0; r < rows; r++)
        {
            var tcElements = new List<XElement>();
            for (int c = 0; c < cols; c++)
            {
                var tc = new XElement(Hp + "tc");

                if (includeCellAddr)
                {
                    tc.Add(new XElement(Hp + "cellAddr",
                        new XAttribute("colAddr", c),
                        new XAttribute("rowAddr", r),
                        new XAttribute("colSpan", 1),
                        new XAttribute("rowSpan", 1)));
                }

                if (includeCellSz)
                {
                    tc.Add(new XElement(Hp + "cellSz",
                        new XAttribute("width", 5000),
                        new XAttribute("height", 1000)));
                }

                if (includeCellMargin)
                {
                    tc.Add(new XElement(Hp + "cellMargin",
                        new XAttribute("left", 100),
                        new XAttribute("right", 100),
                        new XAttribute("top", 50),
                        new XAttribute("bottom", 50)));
                }

                tc.Add(new XElement(Hp + "subList",
                    new XElement(Hp + "p",
                        new XAttribute("paraPrIDRef", "0"),
                        new XAttribute("styleIDRef", "0"),
                        new XElement(Hp + "run",
                            new XAttribute("charPrIDRef", "0"),
                            new XElement(Hp + "t", $"R{r}C{c}")))));

                tcElements.Add(tc);
            }
            trElements.Add(new XElement(Hp + "tr", tcElements));
        }

        var sectionXml = new XDocument(
            new XElement(Hs + "sec",
                new XAttribute(XNamespace.Xmlns + "hs", Hs.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "hp", Hp.NamespaceName),
                new XElement(Hp + "p",
                    new XAttribute("paraPrIDRef", "0"),
                    new XAttribute("styleIDRef", "0"),
                    new XElement(Hp + "run",
                        new XAttribute("charPrIDRef", "0"),
                        new XElement(Hp + "t", "Before table"))),
                new XElement(Hp + "tbl",
                    new XAttribute("rowCnt", rows),
                    new XAttribute("colCnt", cols),
                    trElements)));
        WriteXmlEntry(archive, "Contents/section0.xml", sectionXml);

        return filePath;
    }

    /// <summary>
    /// Create a HWPX with a table cell using tc-attribute format for cellAddr
    /// (legacy format: colAddr/rowAddr directly on hp:tc element).
    /// </summary>
    public static string CreateHwpxWithLegacyCellAddr()
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.hwpx");
        using var stream = File.Create(filePath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteEntry(archive, "mimetype", "application/hwp+zip");

        var containerXml = new XDocument(
            new XElement("container",
                new XAttribute("version", "1.0"),
                new XElement("rootfiles",
                    new XElement("rootfile",
                        new XAttribute("full-path", "Contents/content.hpf"),
                        new XAttribute("media-type", "application/hwpml-package+xml")))));
        WriteXmlEntry(archive, "META-INF/container.xml", containerXml);

        var hpfXml = new XDocument(
            new XElement(Opf + "package",
                new XAttribute(XNamespace.Xmlns + "hpf", Opf.NamespaceName),
                new XElement(Opf + "manifest",
                    new XElement(Opf + "item",
                        new XAttribute("id", "header"),
                        new XAttribute("href", "header.xml"),
                        new XAttribute("media-type", "application/xml")),
                    new XElement(Opf + "item",
                        new XAttribute("id", "section0"),
                        new XAttribute("href", "section0.xml"),
                        new XAttribute("media-type", "application/xml+section"))),
                new XElement(Opf + "spine",
                    new XElement(Opf + "itemref",
                        new XAttribute("idref", "section0")))));
        WriteXmlEntry(archive, "Contents/content.hpf", hpfXml);

        var headerXml = new XDocument(
            new XElement(Hh + "head",
                new XAttribute(XNamespace.Xmlns + "hh", Hh.NamespaceName),
                new XElement(Hh + "charProperties",
                    new XElement(Hh + "charPr",
                        new XAttribute("id", "0"),
                        new XAttribute("height", "1000"),
                        new XElement(Hh + "fontRef",
                            new XAttribute("hangul", "0"),
                            new XAttribute("latin", "0"),
                            new XAttribute("hanja", "0"),
                            new XAttribute("japanese", "0"),
                            new XAttribute("other", "0"),
                            new XAttribute("symbol", "0"),
                            new XAttribute("user", "0")))),
                new XElement(Hh + "paraProperties",
                    new XElement(Hh + "paraPr",
                        new XAttribute("id", "0"),
                        new XElement(Hh + "align",
                            new XAttribute("horizontal", "LEFT"),
                            new XAttribute("vertical", "BASELINE")))),
                new XElement(Hh + "styles",
                    new XAttribute("itemCnt", "1"),
                    new XElement(Hh + "style",
                        new XAttribute("id", "0"),
                        new XAttribute("type", "PARA"),
                        new XAttribute("name", "바탕글"),
                        new XAttribute("engName", "Normal"),
                        new XAttribute("paraPrIDRef", "0"),
                        new XAttribute("charPrIDRef", "0")))));
        WriteXmlEntry(archive, "Contents/header.xml", headerXml);

        // Table with legacy cellAddr (attributes directly on tc, no child element)
        var sectionXml = new XDocument(
            new XElement(Hs + "sec",
                new XAttribute(XNamespace.Xmlns + "hs", Hs.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "hp", Hp.NamespaceName),
                new XElement(Hp + "tbl",
                    new XAttribute("rowCnt", 1),
                    new XAttribute("colCnt", 2),
                    new XElement(Hp + "tr",
                        new XElement(Hp + "tc",
                            new XAttribute("colAddr", 0),
                            new XAttribute("rowAddr", 0),
                            new XAttribute("colSpan", 1),
                            new XAttribute("rowSpan", 1),
                            new XElement(Hp + "cellSz", new XAttribute("width", 5000), new XAttribute("height", 1000)),
                            new XElement(Hp + "cellMargin",
                                new XAttribute("left", 100), new XAttribute("right", 100),
                                new XAttribute("top", 50), new XAttribute("bottom", 50)),
                            new XElement(Hp + "subList",
                                new XElement(Hp + "p",
                                    new XAttribute("paraPrIDRef", "0"),
                                    new XAttribute("styleIDRef", "0"),
                                    new XElement(Hp + "run",
                                        new XAttribute("charPrIDRef", "0"),
                                        new XElement(Hp + "t", "Cell 0,0"))))),
                        new XElement(Hp + "tc",
                            new XAttribute("colAddr", 1),
                            new XAttribute("rowAddr", 0),
                            new XAttribute("colSpan", 1),
                            new XAttribute("rowSpan", 1),
                            new XElement(Hp + "cellSz", new XAttribute("width", 5000), new XAttribute("height", 1000)),
                            new XElement(Hp + "cellMargin",
                                new XAttribute("left", 100), new XAttribute("right", 100),
                                new XAttribute("top", 50), new XAttribute("bottom", 50)),
                            new XElement(Hp + "subList",
                                new XElement(Hp + "p",
                                    new XAttribute("paraPrIDRef", "0"),
                                    new XAttribute("styleIDRef", "0"),
                                    new XElement(Hp + "run",
                                        new XAttribute("charPrIDRef", "0"),
                                        new XElement(Hp + "t", "Cell 1,0")))))))));
        WriteXmlEntry(archive, "Contents/section0.xml", sectionXml);

        return filePath;
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        // R6: Use UTF8 without BOM to avoid corrupting the mimetype entry
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static void WriteXmlEntry(ZipArchive archive, string name, XDocument doc)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        doc.Save(stream);
    }
}
