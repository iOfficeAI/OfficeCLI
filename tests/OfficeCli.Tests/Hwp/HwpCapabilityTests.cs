using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public class HwpCapabilityTests : IDisposable
{
    private readonly string? _oldEngine = Environment.GetEnvironmentVariable("OFFICECLI_HWP_ENGINE");
    private readonly string? _oldBridge = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH");

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", _oldEngine);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", _oldBridge);
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
}
