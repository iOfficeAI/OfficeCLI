using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public class HwpSavePolicySchemaTests
{
    [Theory]
    [InlineData("schemas/interfaces/save-policy.v1.schema.json", "officecli/interfaces/save-policy/v1")]
    [InlineData("schemas/interfaces/save-transaction.v1.schema.json", "officecli/interfaces/save-transaction/v1")]
    public void Phase34SafeSaveSchemasArePresentAndVersioned(string relativePath, string schemaId)
    {
        var root = JsonNode.Parse(File.ReadAllText(LocateRepoFile(relativePath)))!;

        Assert.Equal(schemaId, root["$id"]!.GetValue<string>());
        Assert.Equal(1, root["properties"]!["schemaVersion"]!["const"]!.GetValue<int>());
    }

    [Fact]
    public void SavePolicyRequiresTransactionalValidationAndBackupFields()
    {
        var root = JsonNode.Parse(File.ReadAllText(LocateRepoFile("schemas/interfaces/save-policy.v1.schema.json")))!;
        var required = root["required"]!.AsArray().Select(node => node!.GetValue<string>()).ToArray();

        Assert.Contains("outputRequired", required);
        Assert.Contains("backupRequired", required);
        Assert.Contains("transactionRequired", required);
        Assert.Contains("validationRequired", required);
    }

    [Fact]
    public void SaveTransactionCarriesVerificationEvidence()
    {
        var root = JsonNode.Parse(File.ReadAllText(LocateRepoFile("schemas/interfaces/save-transaction.v1.schema.json")))!;
        var properties = root["properties"]!.AsObject();

        Assert.True(properties.ContainsKey("backupPath"));
        Assert.True(properties.ContainsKey("manifestPath"));
        Assert.True(properties.ContainsKey("semanticDelta"));
        Assert.True(properties.ContainsKey("visualDelta"));
        Assert.True(properties.ContainsKey("packageIntegrity"));
    }

    [Fact]
    public void HwpSavePolicyHelpReferencesSaveTransactionSchema()
    {
        var (exitCode, stdout) = Invoke(["help", "hwp", "save-policy", "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(0, exitCode);
        Assert.Equal("hwp", root["format"]!.GetValue<string>());
        Assert.Equal("save-policy", root["element"]!.GetValue<string>());
        Assert.Equal(
            "schemas/interfaces/save-transaction.v1.schema.json",
            root["properties"]!["transactionSchema"]!["examples"]![0]!.GetValue<string>());
    }

    [Theory]
    [InlineData("hwp")]
    [InlineData("hwpx")]
    public void SchemaValidateStillAcceptsHwpFormatsAfterSafeSaveUpdates(string format)
    {
        var (exitCode, stdout) = Invoke(["schema", "validate", "--format", format, "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(0, exitCode);
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.True(root["data"]!["ok"]!.GetValue<bool>());
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
