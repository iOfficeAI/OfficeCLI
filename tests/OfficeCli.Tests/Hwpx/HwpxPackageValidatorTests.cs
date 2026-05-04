using System.IO.Compression;
using System.Text;
using OfficeCli.Handlers.Hwpx.Validation;

namespace OfficeCli.Tests.Hwpx;

public sealed class HwpxPackageValidatorTests : IDisposable
{
    private readonly List<string> _paths = new();

    public void Dispose()
    {
        foreach (var path in _paths)
            try { File.Delete(path); } catch { }
    }

    [Fact]
    public void ValidHwpxPassesPackageIntegrity()
    {
        var path = Track(HwpxTestHelper.CreateMinimalHwpx("hello"));

        var result = HwpxPackageValidator.Validate(path);

        Assert.Contains(result.Checks, check => check.Name == "zip-open" && check.Ok);
        Assert.Contains(result.Checks, check => check.Name == "xml-well-formed" && check.Ok);
        Assert.Contains(result.Checks, check => check.Name == "package-integrity" && check.Ok);
        Assert.True((int)result.PackageIntegrity["entryCount"]! > 0);
    }

    [Fact]
    public void BrokenZipFailsZipOpen()
    {
        var path = Track(Path.Combine(Path.GetTempPath(), $"broken-{Guid.NewGuid():N}.hwpx"));
        File.WriteAllText(path, "not a zip");

        var result = HwpxPackageValidator.Validate(path);

        Assert.Contains(result.Checks, check => check.Name == "zip-open" && !check.Ok);
        Assert.Contains(result.Checks, check => check.Name == "package-integrity" && !check.Ok);
    }

    [Fact]
    public void MalformedXmlFailsXmlWellFormed()
    {
        var path = Track(Path.Combine(Path.GetTempPath(), $"malformed-{Guid.NewGuid():N}.hwpx"));
        using (var stream = File.Create(path))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            WriteEntry(archive, "mimetype", "application/hwp+zip");
            WriteEntry(archive, "META-INF/container.xml", "<container><rootfiles></container>");
            WriteEntry(archive, "Contents/content.hpf", "<package>");
            WriteEntry(archive, "Contents/header.xml", "<head/>");
            WriteEntry(archive, "Contents/section0.xml", "<section/>");
        }

        var result = HwpxPackageValidator.Validate(path);

        Assert.Contains(result.Checks, check => check.Name == "xml-well-formed" && !check.Ok);
        Assert.Contains(result.Checks, check => check.Name == "package-integrity" && !check.Ok);
    }

    [Fact]
    public void MissingBinDataTargetFailsReferenceCheck()
    {
        var path = Track(Path.Combine(Path.GetTempPath(), $"missing-bin-{Guid.NewGuid():N}.hwpx"));
        using (var stream = File.Create(path))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            WriteEntry(archive, "mimetype", "application/hwp+zip");
            WriteEntry(archive, "META-INF/container.xml", """
<container><rootfiles><rootfile full-path="Contents/content.hpf"/></rootfiles></container>
""");
            WriteEntry(archive, "Contents/content.hpf", """
<opf:package xmlns:opf="http://www.idpf.org/2007/opf"><opf:manifest><opf:item id="header" href="header.xml"/><opf:item id="section0" href="section0.xml"/></opf:manifest></opf:package>
""");
            WriteEntry(archive, "Contents/header.xml", """
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"/>
""");
            WriteEntry(archive, "Contents/section0.xml", """
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p><hp:run><hp:pic binaryItemIDRef="missing-image"/></hp:run></hp:p></hs:sec>
""");
        }

        var result = HwpxPackageValidator.Validate(path);

        Assert.Contains(result.Checks, check => check.Name == "bindata-references-present" && !check.Ok);
        Assert.Contains(result.Checks, check => check.Name == "package-integrity" && !check.Ok);
    }

    private string Track(string path)
    {
        _paths.Add(path);
        return path;
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
