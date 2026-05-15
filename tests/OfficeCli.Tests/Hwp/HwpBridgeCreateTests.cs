using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void CreateBlank_DelegatesToRhwpApiBridgeAndCreatesOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var output = CreateOutput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            ["create-blank", "--output", output, "--json"],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.True(root["created"]!.GetValue<bool>());
        Assert.Equal(output, root["output"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliCreateHwp_RoutesThroughRhwpApiBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(["create", output, "--json"]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
    }
}
