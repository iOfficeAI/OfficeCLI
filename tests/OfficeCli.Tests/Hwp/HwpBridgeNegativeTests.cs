using System.Text.Json.Nodes;
using OfficeCli;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public class HwpBridgeNegativeTests : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly string? _oldEngine = Environment.GetEnvironmentVariable("OFFICECLI_HWP_ENGINE");
    private readonly string? _oldBridge = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH");

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", _oldEngine);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", _oldBridge);
        foreach (var path in _tempFiles)
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void HwpViewTextJson_WithoutExperimentalEnv_ReturnsBridgeNotEnabled()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", null);
        var hwp = CreateFakeHwp();

        var (exitCode, stdout) = Invoke(["view", hwp, "text", "--json"]);

        Assert.Equal(1, exitCode);
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("bridge_not_enabled", root["error"]!["code"]!.GetValue<string>());
        Assert.Equal("hwp", root["error"]!["format"]!.GetValue<string>());
        Assert.Equal("read_text", root["error"]!["operation"]!.GetValue<string>());
        Assert.Equal("none", root["error"]!["engine"]!.GetValue<string>());
        Assert.Equal("none", root["error"]!["engineMode"]!.GetValue<string>());
    }

    [Fact]
    public void HwpViewTextJson_WithExperimentalEnvButMissingBridge_ReturnsBridgeMissing()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", "/tmp/missing-rhwp-officecli-bridge");
        var hwp = CreateFakeHwp();

        var (exitCode, stdout) = Invoke(["view", hwp, "text", "--json"]);

        Assert.Equal(1, exitCode);
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("bridge_missing", root["error"]!["code"]!.GetValue<string>());
        Assert.Equal("hwp", root["error"]!["format"]!.GetValue<string>());
        Assert.Equal("read_text", root["error"]!["operation"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["error"]!["engine"]!.GetValue<string>());
        Assert.Equal("experimental", root["error"]!["engineMode"]!.GetValue<string>());
    }

    [Fact]
    public async Task CustomEngine_BinaryHwpFillField_ReturnsMutationForbidden()
    {
        var engine = new CustomHwpxEngine();
        var request = new HwpFillFieldRequest(
            HwpFormat.Hwp,
            CreateFakeHwp(),
            CreateFakeOutput(".hwp"),
            new Dictionary<string, string> { ["name"] = "value" },
            true);

        var ex = await Assert.ThrowsAsync<HwpEngineException>(
            () => engine.FillFieldAsync(request, CancellationToken.None));

        Assert.Equal("binary_hwp_mutation_forbidden", ex.Error.Code);
        Assert.Equal("hwp", ex.Error.Format);
        Assert.Equal("fill_field", ex.Error.Operation);
    }

    [Fact]
    public async Task RhwpBridgeEngine_BinaryHwpSaveOriginal_ReturnsWriteForbidden()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", CreateFakeBridge());
        var engine = HwpEngineSelector.GetEngine();
        var request = new HwpSaveOriginalRequest(
            HwpFormat.Hwp,
            CreateFakeHwp(),
            CreateFakeOutput(".hwp"),
            true);

        var ex = await Assert.ThrowsAsync<HwpEngineException>(
            () => engine.SaveOriginalAsync(request, CancellationToken.None));

        Assert.Equal("binary_hwp_write_forbidden", ex.Error.Code);
        Assert.Equal("hwp", ex.Error.Format);
        Assert.Equal("save_original", ex.Error.Operation);
    }

    private string CreateFakeHwp()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_fake_{Guid.NewGuid():N}.hwp");
        File.WriteAllBytes(path, [0x48, 0x57, 0x50]);
        _tempFiles.Add(path);
        return path;
    }

    private string CreateFakeBridge()
    {
        var path = Path.Combine(Path.GetTempPath(), $"rhwp-officecli-bridge-{Guid.NewGuid():N}");
        File.WriteAllText(path, "#!/bin/sh\nexit 0\n");
        _tempFiles.Add(path);
        return path;
    }

    private string CreateFakeOutput(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli_out_{Guid.NewGuid():N}{extension}");
        _tempFiles.Add(path);
        return path;
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
