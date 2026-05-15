using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void ReadText_DelegatesToRhwpExportTextAndReturnsCanonicalJson()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeRhwp = CreateFakeRhwp();
        var input = CreateInput(".hwp");

        var result = RunBridge(bridgeDll, fakeRhwp,
            ["read-text", "--format", "hwp", "--input", input, "--json"]);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("before before", root["text"]!.GetValue<string>());
        Assert.Equal("rhwp v0.test", root["engineVersion"]!.GetValue<string>());
        Assert.Equal("hwp", root["format"]!.GetValue<string>());
        Assert.Equal(1, root["pages"]![0]!["page"]!.GetValue<int>());
        Assert.Equal("before before", root["pages"]![0]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void RenderSvg_DelegatesToRhwpExportSvgAndReturnsManifestJson()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeRhwp = CreateFakeRhwp();
        var input = CreateInput(".hwp");
        var outDir = CreateDirectory();

        var result = RunBridge(bridgeDll, fakeRhwp,
            ["render-svg", "--format", "hwp", "--input", input, "--out-dir", outDir, "--page", "1", "--json"]);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("rhwp v0.test", root["engineVersion"]!.GetValue<string>());
        Assert.Equal("hwp", root["format"]!.GetValue<string>());
        Assert.Equal(Path.Combine(outDir, "manifest.json"), root["manifest"]!.GetValue<string>());
        Assert.Equal(1, root["pages"]![0]!["page"]!.GetValue<int>());
        Assert.EndsWith("page.svg", root["pages"]![0]!["path"]!.GetValue<string>());
        Assert.True(File.Exists(root["pages"]![0]!["path"]!.GetValue<string>()));
        Assert.Equal(64, root["pages"]![0]!["sha256"]!.GetValue<string>().Length);
    }

    [Fact]
    public void OfficeCliViewText_CanRunBridgeDllViaDotnet()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable(
            "OFFICECLI_RHWP_API_BIN",
            Path.Combine(Path.GetTempPath(), "officecli-test-no-rhwp-api"));
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, "text", "--json"]);

        Assert.True(exitCode == 0, stdout);
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal("before before", root["data"]!["text"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
        Assert.Equal("rhwp v0.test", root["data"]!["engineVersion"]!.GetValue<string>());
    }

    [Fact]
    public void ReadText_PrefersRhwpApiBridgeWhenConfigured()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            ["read-text", "--format", "hwp", "--input", input, "--json"],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("before before", root["text"]!.GetValue<string>());
        Assert.Equal("rhwp-api v0.test", root["engineVersion"]!.GetValue<string>());
    }
}
