using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;
using OfficeCli.Handlers.Hwp.SafeSave;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public sealed class HwpSafeSaveManifestTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"officecli-safe-save-manifest-{Guid.NewGuid():N}");

    public HwpSafeSaveManifestTests()
    {
        Directory.CreateDirectory(_root);
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    [Fact]
    public async Task OutputModeWritesManifestAfterSuccessfulTransaction()
    {
        var input = CreateFile("input/source.hwp", "source");
        var output = Path.Combine(_root, "out/result.hwp");

        var transaction = await RunAsync(input, output, [new SafeSaveCheck("provider-readback", true, "info")]);

        Assert.True(transaction.Ok);
        Assert.NotNull(transaction.ManifestPath);
        Assert.True(File.Exists(transaction.ManifestPath));
        var root = JsonNode.Parse(File.ReadAllText(transaction.ManifestPath!))!;
        Assert.True(root["ok"]!.GetValue<bool>());
        Assert.Equal("output", root["mode"]!.GetValue<string>());
        Assert.Equal(transaction.ManifestPath, root["manifestPath"]!.GetValue<string>());
        Assert.Contains(
            root["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "manifest-write" && check["ok"]!.GetValue<bool>());
    }

    [Fact]
    public async Task ValidationFailureWritesManifestWithFailedCheck()
    {
        var input = CreateFile("input/source.hwp", "source");
        var output = Path.Combine(_root, "out/result.hwp");

        var transaction = await RunAsync(input, output, [new SafeSaveCheck("provider-readback", false, "error", "readback failed")]);

        Assert.False(transaction.Ok);
        Assert.False(File.Exists(output));
        Assert.NotNull(transaction.ManifestPath);
        var root = JsonNode.Parse(File.ReadAllText(transaction.ManifestPath!))!;
        Assert.False(root["ok"]!.GetValue<bool>());
        Assert.Contains(
            root["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "provider-readback" && !check["ok"]!.GetValue<bool>());
    }

    [Fact]
    public void VisualValidatorRecordsSvgPageCountAndSha()
    {
        var result = SafeSaveVisualValidator.FromRenderResult(new HwpRenderResult(
            [new HwpRenderedPage(1, "/tmp/page.svg", "abc123")],
            "/tmp/manifest.json",
            "rhwp-bridge",
            "rhwp v0.test",
            [],
            []));

        Assert.Contains(result.Checks, check => check.Name == "visual-render" && check.Ok);
        Assert.Equal(1, result.VisualDelta!["pageCount"]);
        Assert.Equal("abc123", result.VisualDelta!["firstPageSha256"]);
    }

    private async Task<SafeSaveTransaction> RunAsync(
        string input,
        string output,
        IReadOnlyList<SafeSaveCheck> validationChecks)
    {
        var runner = new SafeSaveRunner();
        return await runner.RunAsync(
            Options(input, output, SafeSavePolicy.OutputMode("temp-write", "provider-readback")),
            temp =>
            {
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks(validationChecks)),
            CancellationToken.None);
    }

    private static SafeSaveOptions Options(string input, string output, SafeSavePolicy policy) => new(
        input,
        output,
        InPlace: false,
        Backup: false,
        Verify: true,
        HwpCapabilityConstants.OperationReplaceText,
        HwpCapabilityConstants.FormatHwp,
        policy);

    private string CreateFile(string relativePath, string content)
    {
        var path = Path.Combine(_root, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
        return path;
    }
}
