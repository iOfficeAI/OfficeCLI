using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public class HwpSchemaHelpTests
{
    [Theory]
    [InlineData("hwp", "table-cell")]
    [InlineData("hwp", "provider-rhwp")]
    [InlineData("hwpx", "paragraph")]
    [InlineData("hwpx", "raw-xml")]
    [InlineData("hwpx", "provider-custom")]
    public void Help_LoadsHwpAndHwpxSchemas(string format, string element)
    {
        var (exitCode, stdout) = Invoke(["help", format, element, "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(0, exitCode);
        Assert.Equal(format, root["format"]!.GetValue<string>());
        Assert.Equal(element, root["element"]!.GetValue<string>());
    }

    [Fact]
    public void Help_ListsHwpAndHwpxFormats()
    {
        var (exitCode, stdout) = Invoke(["help"]);

        Assert.Equal(0, exitCode);
        Assert.Contains("hwpx", stdout);
        Assert.Contains("hwp", stdout);
    }

    [Theory]
    [InlineData("hwp", "table-map")]
    [InlineData("hwp", "validate-output")]
    [InlineData("hwpx", "table-cell")]
    [InlineData("hwpx", "render-svg")]
    public void SchemaList_IncludesNewHwpEntries(string format, string element)
    {
        var (exitCode, stdout) = Invoke(["schema", "list", "--format", format, "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(0, exitCode);
        Assert.True(root["success"]!.GetValue<bool>());
        var entries = root["data"]!["entries"]!.AsArray();
        Assert.Contains(entries, node =>
            node!["format"]!.GetValue<string>() == format
            && node["element"]!.GetValue<string>() == element);
    }

    [Theory]
    [InlineData("hwp")]
    [InlineData("hwpx")]
    public void SchemaValidate_ParsesNewHwpSchemas(string format)
    {
        var (exitCode, stdout) = Invoke(["schema", "validate", "--format", format, "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(0, exitCode);
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.True(root["data"]!["ok"]!.GetValue<bool>());
        Assert.True(root["data"]!["checked"]!.GetValue<int>() > 0);
    }

    [Theory]
    [InlineData("schemas/interfaces/capability-result.v1.schema.json")]
    [InlineData("schemas/interfaces/edit-result.v1.schema.json")]
    [InlineData("schemas/interfaces/rhwp-sidecar-request.v1.schema.json")]
    [InlineData("schemas/interfaces/rhwp-sidecar-response.v1.schema.json")]
    [InlineData("docs/providers/rhwp-sidecar-contract.md")]
    public void Phase30And33ContractFilesExist(string relativePath)
    {
        Assert.True(File.Exists(LocateRepoFile(relativePath)), relativePath);
    }

    private static string LocateRepoFile(string relativePath)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, relativePath);
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new FileNotFoundException($"Required test file was not found: {relativePath}");
    }

    private static (int ExitCode, string Stdout) Invoke(string[] args)
    {
        var root = CommandBuilder.BuildRootCommand();
        var originalOut = Console.Out;
        using var writer = new StringWriter();
        Console.SetOut(writer);
        try
        {
            var exitCode = root.Parse(args).Invoke();
            return (exitCode, writer.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
        }
    }
}
