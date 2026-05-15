using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void ConvertToEditable_DelegatesToRhwpApiBridgeAndCreatesOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            ["convert-to-editable", "--format", "hwp", "--input", input, "--output", output, "--json"],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.True(root["converted"]!["converted"]!.GetValue<bool>());
        Assert.Equal(output, root["output"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliSetConvertToEditable_RoutesThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/convert-to-editable",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal(output, root["data"]!["outputPath"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }
}
