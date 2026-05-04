using System.Diagnostics;
using System.Text.Json.Nodes;
using OfficeCli;
using OfficeCli.Handlers.Hwp;
using OfficeCli.Tests.Hwpx;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public partial class HwpBridgeSidecarTests : IDisposable
{
    private readonly List<string> _tempPaths = new();
    private readonly string? _oldEngine = Environment.GetEnvironmentVariable("OFFICECLI_HWP_ENGINE");
    private readonly string? _oldBridge = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH");
    private readonly string? _oldRhwp = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BIN");
    private readonly string? _oldRhwpApi = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_API_BIN");

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", _oldEngine);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", _oldBridge);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", _oldRhwp);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", _oldRhwpApi);
        foreach (var path in _tempPaths)
        {
            try
            {
                if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
                else File.Delete(path);
            }
            catch { }
        }
    }

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
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, "text", "--json"]);

        Assert.Equal(0, exitCode);
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal("before before", root["data"]!["text"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
        Assert.Equal("rhwp v0.test", root["data"]!["engineVersion"]!.GetValue<string>());
    }

    [Fact]
    public void ListFields_DelegatesToRhwpApiBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            ["list-fields", "--format", "hwp", "--input", input, "--json"],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("rhwp-api v0.test", root["engineVersion"]!.GetValue<string>());
        Assert.Equal("hwp", root["format"]!.GetValue<string>());
        Assert.Equal(7, root["fields"]![0]!["fieldId"]!.GetValue<int>());
        Assert.Equal("applicant", root["fields"]![0]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void GetField_DelegatesToRhwpApiBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            ["get-field", "--format", "hwp", "--input", input, "--name", "applicant", "--json"],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("홍길동", root["field"]!["value"]!.GetValue<string>());
        Assert.Equal(7, root["field"]!["fieldId"]!.GetValue<int>());
    }

    [Fact]
    public void SetField_DelegatesToRhwpApiBridgeAndCreatesOutput()
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
                "set-field", "--format", "hwp", "--input", input, "--output", output,
                "--name", "applicant", "--value", "김철수", "--json"
            ],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal(output, root["output"]!.GetValue<string>());
        Assert.Equal("김철수", root["field"]!["newValue"]!.GetValue<string>());
    }

    [Fact]
    public void ReplaceText_DelegatesToRhwpApiBridgeAndCreatesOutput()
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
                "replace-text", "--format", "hwp", "--input", input, "--output", output,
                "--query", "before", "--value", "after", "--mode", "all", "--json"
            ],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal(output, root["output"]!.GetValue<string>());
        Assert.Equal(2, root["replacement"]!["count"]!.GetValue<int>());
    }

    [Fact]
    public void GetCellText_DelegatesToRhwpApiBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwp");

        var result = RunBridge(
            bridgeDll,
            CreateFakeRhwp(),
            [
                "get-cell-text", "--format", "hwp", "--input", input,
                "--section", "0", "--parent-para", "2", "--control", "0",
                "--cell", "1", "--cell-para", "0", "--json"
            ],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("셀값", root["cell"]!["text"]!.GetValue<string>());
    }


    [Fact]
    public async Task RhwpBridgeEngineFillField_CallsSetFieldAndReturnsMutationEvidence()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var engine = HwpEngineSelector.GetEngine();
        var result = await engine.FillFieldAsync(
            new HwpFillFieldRequest(
                HwpFormat.Hwp,
                input,
                output,
                new Dictionary<string, string> { ["applicant"] = "김철수" },
                true),
            CancellationToken.None);

        Assert.Equal(output, result.OutputPath);
        Assert.Equal("rhwp-bridge", result.Engine);
        Assert.Equal("rhwp-api v0.test", result.EngineVersion);
        Assert.True(File.Exists(output));
        Assert.Contains(result.Warnings, w => w.Contains("experimental", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void OfficeCliSetField_RoutesBinaryHwpThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/field",
                "--prop", "name=applicant",
                "--prop", "value=김철수",
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

    [Fact]
    public void OfficeCliSetField_CanUseFieldId()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/field",
                "--prop", "id=7",
                "--prop", "value=김철수",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Contains("#7", root["message"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliSetText_RoutesBinaryHwpThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=before",
                "--prop", "value=after",
                "--prop", "mode=all",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal(output, root["data"]!["outputPath"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
        Assert.Equal("output", root["data"]!["transaction"]!["mode"]!.GetValue<string>());
        Assert.True(root["data"]!["transaction"]!["verified"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "provider-readback" && check["ok"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "semantic-delta" && check["ok"]!.GetValue<bool>());
        Assert.Equal(2, root["data"]!["transaction"]!["semanticDelta"]!["sourceOldCount"]!.GetValue<int>());
        Assert.Equal(0, root["data"]!["transaction"]!["semanticDelta"]!["outputOldCount"]!.GetValue<int>());
    }

    [Fact]
    public void OfficeCliSetText_RoutesExperimentalHwpxThroughPackageIntegrityCheck()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = HwpxTestHelper.CreateMinimalHwpx("before before");
        _tempPaths.Add(input);
        var output = CreateOutput(".hwpx");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=before",
                "--prop", "value=after",
                "--prop", "mode=all",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.True(exitCode == 0, stdout);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        var transaction = root["data"]!["transaction"]!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.True(transaction["verified"]!.GetValue<bool>());
        Assert.Contains(
            transaction["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "package-integrity" && check["ok"]!.GetValue<bool>());
        Assert.NotNull(transaction["packageIntegrity"]);
        Assert.True(transaction["packageIntegrity"]!["entryCount"]!.GetValue<int>() > 0);
    }

    [Fact]
    public void OfficeCliSetText_InPlaceReturnsSafeSaveTransactionError()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=before",
                "--prop", "value=after",
                "--in-place",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.Equal("before before", File.ReadAllText(input));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("in-place", root["data"]!["transaction"]!["mode"]!.GetValue<string>());
        Assert.False(root["data"]!["transaction"]!["ok"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "in-place-not-ready");
    }

    [Fact]
    public void OfficeCliSetText_SemanticFailureDoesNotPublishOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=before",
                "--prop", "value=before",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.False(File.Exists(output));
        Assert.Equal("before before", File.ReadAllText(input));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("output", root["data"]!["transaction"]!["mode"]!.GetValue<string>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "semantic-delta" && !check["ok"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "source-preserved" && check["ok"]!.GetValue<bool>());
    }

}
