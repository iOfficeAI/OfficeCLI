using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public class HwpCapabilityTests : IDisposable
{
    private readonly string? _oldEngine = Environment.GetEnvironmentVariable("OFFICECLI_HWP_ENGINE");
    private readonly string? _oldBridge = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH");
    private readonly string? _oldRhwp = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BIN");
    private readonly string? _oldApi = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_API_BIN");

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", _oldEngine);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", _oldBridge);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", _oldRhwp);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", _oldApi);
    }

    [Fact]
    public void Report_ContainsCanonicalFormatsAndOperations()
    {
        var report = HwpCapabilityFactory.BuildReport("officecli:test");

        Assert.Equal(2, report.SchemaVersion);
        Assert.Contains("hwpx", report.Formats.Keys);
        Assert.Contains("hwp", report.Formats.Keys);

        foreach (var capability in report.Formats.Values)
        {
            Assert.Contains("read_text", capability.Operations.Keys);
            Assert.Contains("render_svg", capability.Operations.Keys);
            Assert.Contains("fill_field", capability.Operations.Keys);
            Assert.Contains("save_original", capability.Operations.Keys);
            Assert.Contains("save_as_hwp", capability.Operations.Keys);
        }
    }

    [Fact]
    public void Json_IncludesNullEngineVersionForUnsupportedOperations()
    {
        var report = HwpCapabilityFactory.BuildReport("officecli:test");
        var envelope = HwpCapabilityJsonMapper.BuildEnvelope(report);
        var json = envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions);

        Assert.Contains("\"success\": true", json);
        Assert.Contains("\"warnings\": []", json);
        Assert.Contains("\"engineVersion\": null", json);
        Assert.Contains("\"unsupportedReason\": \"bridge_not_enabled\"", json);
        Assert.DoesNotContain("RoundTripVerified", json);
        Assert.DoesNotContain("\"error\": null", json);
    }

    [Fact]
    public void BinaryHwp_MutationIsExperimentalButNotReadyWithoutBridge()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", null);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", null);
        var report = HwpCapabilityFactory.BuildReport("officecli:test");
        var hwp = report.Formats["hwp"];

        Assert.Equal("unsupported", hwp.ReadStatus);
        Assert.Equal("unsupported", hwp.WriteStatus);
        Assert.Equal("bridge_not_enabled",
            hwp.Operations["fill_field"].UnsupportedReason);
        Assert.Equal(HwpOperationStatus.Experimental,
            hwp.Operations["fill_field"].Status);
        Assert.Equal("binary_hwp_write_forbidden",
            hwp.Operations["save_original"].UnsupportedReason);
        Assert.Equal("binary_hwp_write_forbidden",
            hwp.Operations["save_as_hwp"].UnsupportedReason);
    }

    [Fact]
    public void HwpHelpJson_IncludesAgentRecipesAndDoctor()
    {
        var (exitCode, stdout) = Invoke(["hwp", "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(0, exitCode);
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal(1, root["data"]!["schemaVersion"]!.GetValue<int>());
        Assert.Equal("officecli hwp doctor --json", root["data"]!["doctor"]!.GetValue<string>());
        Assert.Equal(
            "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json",
            root["data"]!["recipes"]!["replaceText"]!.GetValue<string>());
    }

    [Fact]
    public void HwpDoctorJson_ReportsMissingRuntimeAndNextCommand()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", null);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", null);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", null);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", null);

        var (exitCode, stdout) = Invoke(["hwp", "doctor", "--json"]);
        var root = JsonNode.Parse(stdout)!;

        Assert.Equal(2, exitCode);
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.False(root["data"]!["ok"]!.GetValue<bool>());
        Assert.Equal("officecli help hwp", root["data"]!["nextCommand"]!.GetValue<string>());
        Assert.Equal("OFFICECLI_HWP_ENGINE",
            root["data"]!["checks"]![0]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void HwpBridgeErrorJson_IncludesNextCommand()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", null);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", null);
        var input = Path.Combine(Path.GetTempPath(), $"officecli-hwp-{Guid.NewGuid():N}.hwp");
        File.WriteAllText(input, "not a real hwp");

        try
        {
            var (exitCode, stdout) = Invoke(["view", input, "text", "--json"]);
            var root = JsonNode.Parse(stdout)!;

            Assert.Equal(1, exitCode);
            Assert.False(root["success"]!.GetValue<bool>());
            Assert.Equal("bridge_not_enabled", root["error"]!["code"]!.GetValue<string>());
            Assert.Equal("officecli hwp doctor --json", root["error"]!["nextCommand"]!.GetValue<string>());
        }
        finally
        {
            File.Delete(input);
        }
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
