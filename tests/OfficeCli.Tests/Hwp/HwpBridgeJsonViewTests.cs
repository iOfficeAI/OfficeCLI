using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Theory]
    [InlineData("markdown", "markdown", "Fake HWP")]
    [InlineData("info", "info", "pages")]
    [InlineData("diagnostics", "diagnostics", "")]
    [InlineData("dump", "dump", "full control dump")]
    [InlineData("pages", "dump", "page 1 dump")]
    [InlineData("tables", "cells", "셀값")]
    public void OfficeCliViewHwpJsonModes_RouteThroughRhwpBridge(string mode, string expectedKey, string expectedText)
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, mode, "--json"]);

        Assert.Equal(0, exitCode);
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
        Assert.NotNull(root["data"]![expectedKey]);
        if (expectedKey == "cells")
            Assert.Equal(expectedText, root["data"]![expectedKey]![0]!["text"]!.GetValue<string>());
        else if (!string.IsNullOrEmpty(expectedText))
            Assert.Contains(expectedText, root["data"]![expectedKey]!.ToJsonString());
    }

    [Fact]
    public void OfficeCliViewHwpPdf_RoutesThroughRhwpBridgeAndCreatesPdf()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".pdf");

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, "pdf", "--page", "1", "--out", output, "--json"]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.Equal(output, root["data"]!["pdf"]!["path"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewHwpxPdf_UsesInstalledRhwpRuntimeWithoutEngineEnv()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", null);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwpx");
        var output = CreateOutput(".pdf");

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, "pdf", "--page", "1", "--out", output, "--json"]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.Equal(output, root["data"]!["pdf"]!["path"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewHwpPng_RoutesThroughRhwpBridgeAndCreatesPageImage()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var outDir = CreateDirectory();

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, "png", "--page", "1", "--out", outDir, "--json"]);

        Assert.Equal(0, exitCode);
        var root = JsonNode.Parse(stdout)!;
        var pagePath = root["data"]!["pages"]![0]!["path"]!.GetValue<string>();
        Assert.True(File.Exists(pagePath));
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewHwpThumbnail_RoutesThroughRhwpBridgeAndCreatesImage()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".png");

        var (exitCode, stdout) = InvokeOfficeCli(["view", input, "thumbnail", "--out", output, "--json"]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.Equal(output, root["data"]!["thumbnail"]!["path"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliViewHwpTableCell_RoutesThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "view", input, "table-cell",
                "--section", "0",
                "--parent-para", "2",
                "--control", "0",
                "--cell", "1",
                "--cell-para", "0",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        var root = JsonNode.Parse(stdout)!;
        Assert.Equal("셀값", root["data"]!["cell"]!["text"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", root["data"]!["engine"]!.GetValue<string>());
    }
}
