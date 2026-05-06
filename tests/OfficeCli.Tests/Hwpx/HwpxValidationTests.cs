using System.IO.Compression;
using System.Text;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Hwpx;

public sealed class HwpxValidationTests : IDisposable
{
    private readonly List<string> _paths = [];

    public void Dispose()
    {
        foreach (var path in _paths)
            try { File.Delete(path); } catch { }
    }

    [Fact]
    public void ValidateAcceptsRootVersionXml()
    {
        var path = Track(CreatePackage(rootVersionXml: true, includeManifestImage: false));

        using var handler = new HwpxHandler(path, editable: false);
        var errors = handler.Validate();

        Assert.DoesNotContain(errors, e => e.ErrorType == "package_version_missing");
    }

    [Fact]
    public void ValidateTreatsContentHpfBinDataItemsAsReferenced()
    {
        var path = Track(CreatePackage(rootVersionXml: true, includeManifestImage: true));

        using var handler = new HwpxHandler(path, editable: false);
        var errors = handler.Validate();

        Assert.DoesNotContain(errors, e => e.ErrorType == "bindata_orphan");
    }

    [Fact]
    public void ValidateIgnoresZipDirectoryEntriesInBinData()
    {
        var path = Track(CreatePackage(rootVersionXml: true, includeManifestImage: true, includeBinDataDirectory: true));

        using var handler = new HwpxHandler(path, editable: false);
        var errors = handler.Validate();

        Assert.DoesNotContain(errors, e => e.ErrorType == "bindata_orphan" && e.Description.Contains("''"));
    }

    private string Track(string path)
    {
        _paths.Add(path);
        return path;
    }

    private static string CreatePackage(
        bool rootVersionXml,
        bool includeManifestImage,
        bool includeBinDataDirectory = false)
    {
        var path = Path.Combine(Path.GetTempPath(), $"hwpx-validate-{Guid.NewGuid():N}.hwpx");
        using var stream = File.Create(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteEntry(archive, "mimetype", "application/hwp+zip", CompressionLevel.NoCompression);
        WriteEntry(archive, "META-INF/container.xml", """
<container><rootfiles><rootfile full-path="Contents/content.hpf"/></rootfiles></container>
""");
        WriteEntry(archive, "Contents/content.hpf", includeManifestImage ? """
<opf:package xmlns:opf="http://www.idpf.org/2007/opf"><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/><opf:item id="image1" href="BinData/image1.png" media-type="image/png"/></opf:manifest><opf:spine><opf:itemref idref="section0"/></opf:spine></opf:package>
""" : """
<opf:package xmlns:opf="http://www.idpf.org/2007/opf"><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="section0"/></opf:spine></opf:package>
""");
        WriteEntry(archive, "Contents/header.xml", """
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"><hh:charProperties><hh:charPr id="0"/></hh:charProperties><hh:paraProperties><hh:paraPr id="0"/></hh:paraProperties><hh:styles><hh:style id="0"/></hh:styles></hh:head>
""");
        WriteEntry(archive, "Contents/section0.xml", """
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>hello</hp:t></hp:run></hp:p></hs:sec>
""");

        if (rootVersionXml)
            WriteEntry(archive, "version.xml", """
<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" major="5"/>
""");
        else
            WriteEntry(archive, "Contents/version.xml", """
<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" major="5"/>
""");

        if (includeManifestImage)
        {
            if (includeBinDataDirectory)
                archive.CreateEntry("BinData/");
            WriteEntry(archive, "BinData/image1.png", "fake-png");
        }

        return path;
    }

    private static void WriteEntry(
        ZipArchive archive,
        string name,
        string content,
        CompressionLevel compressionLevel = CompressionLevel.Optimal)
    {
        var entry = archive.CreateEntry(name, compressionLevel);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
