using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public class HwpCapabilityTests
{
    [Fact]
    public void Report_ContainsCanonicalFormatsAndOperations()
    {
        var report = HwpCapabilityFactory.BuildReport("officecli:test");

        Assert.Equal(1, report.SchemaVersion);
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
    public void BinaryHwp_MutationAndWriteRemainUnsupported()
    {
        var report = HwpCapabilityFactory.BuildReport("officecli:test");
        var hwp = report.Formats["hwp"];

        Assert.Equal("unsupported", hwp.ReadStatus);
        Assert.Equal("unsupported", hwp.WriteStatus);
        Assert.Equal("binary_hwp_mutation_forbidden",
            hwp.Operations["fill_field"].UnsupportedReason);
        Assert.Equal("binary_hwp_write_forbidden",
            hwp.Operations["save_original"].UnsupportedReason);
        Assert.Equal("binary_hwp_write_forbidden",
            hwp.Operations["save_as_hwp"].UnsupportedReason);
    }
}
