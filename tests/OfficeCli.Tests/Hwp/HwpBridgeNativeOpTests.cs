using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void NativeOp_DelegatesToRhwpApiBridgeAndCreatesOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            ["native-op", "--format", "hwp", "--input", input, "--op", "split-paragraph", "--output", output, "--json"],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("split-paragraph", root["operation"]!.GetValue<string>());
        Assert.Equal(output, root["output"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewNative_RoutesThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            ["view", input, "native", "--op", "get-style-list", "--json"]);

        Assert.Equal(0, exitCode);
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal("get-style-list", root["data"]!["operation"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewNative_RejectsMutatingNativeOp()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "view", input, "native",
                "--op", "split-paragraph",
                "--native-arg", "paragraph=0",
                "--native-arg", "offset=5",
                "--native-arg", $"output={output}",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.False(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Contains("not read-only", root["error"]!["error"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewNative_RejectsOutputArgForReadOp()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "view", input, "native",
                "--op", "get-style-list",
                "--native-arg", $"output={output}",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.False(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Contains("does not accept output paths", root["error"]!["error"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliSetNativeOp_RoutesThroughRhwpBridgeAndCreatesOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/native-op",
                "--prop", "op=split-paragraph",
                "--prop", "paragraph=0",
                "--prop", "offset=5",
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
