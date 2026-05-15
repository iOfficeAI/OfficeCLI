using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void InsertText_DelegatesToRhwpApiBridgeAndCreatesOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            [
                "insert-text", "--format", "hwp", "--input", input,
                "--output", output, "--section", "0", "--paragraph", "0",
                "--offset", "0", "--value", "본문", "--json"
            ],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        Assert.Equal("본문", File.ReadAllText(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.True(root["inserted"]!.GetValue<bool>());
        Assert.Equal(output, root["output"]!.GetValue<string>());
        Assert.Equal("rhwp-api v0.test", root["engineVersion"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliAddHwpText_RoutesThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "add", input, "/text", "--type", "paragraph",
                "--prop", "value=본문",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        Assert.Equal("본문", File.ReadAllText(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal(output, root["data"]!["outputPath"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
        Assert.Equal("output", root["data"]!["transaction"]!["mode"]!.GetValue<string>());
        Assert.True(root["data"]!["transaction"]!["verified"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "source-preserved" && check["ok"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["evidence"]!.AsArray(),
            value => value!.GetValue<string>().Contains("insert-text", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void OfficeCliAddHwpText_RejectsSamePathOutputAndPreservesSource()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var before = File.ReadAllText(input);

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "add", input, "/text", "--type", "paragraph",
                "--prop", "value=본문",
                "--prop", $"output={input}",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.Equal(before, File.ReadAllText(input));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("fixture_validation_failed", root["error"]!["code"]!.GetValue<string>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "same-path-output" && !check["ok"]!.GetValue<bool>());
    }
}
