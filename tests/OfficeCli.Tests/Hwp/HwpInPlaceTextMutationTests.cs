using System.Security.Cryptography;
using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    [Fact]
    public void OfficeCliSetText_InPlaceRequiresVerifyBeforeMutation()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        var input = CreateInput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=before",
                "--prop", "value=after",
                "--in-place",
                "--backup",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.Equal("before before", File.ReadAllText(input));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("hwp_in_place_requires_verify", root["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliSetText_InPlaceSemanticFailureLeavesSourceHashUnchanged()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var beforeHash = Sha256(input);

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/text",
                "--prop", "find=before",
                "--prop", "value=before",
                "--in-place",
                "--backup",
                "--verify",
                "--json"
            ]);

        Assert.Equal(1, exitCode);
        Assert.Equal(beforeHash, Sha256(input));
        Assert.Equal("before before", File.ReadAllText(input));
        var root = JsonNode.Parse(stdout)!;
        var transaction = root["data"]!["transaction"]!;
        Assert.Equal("in-place", transaction["mode"]!.GetValue<string>());
        Assert.False(transaction["ok"]!.GetValue<bool>());
        Assert.Contains(
            transaction["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "semantic-delta" && !check["ok"]!.GetValue<bool>());
        Assert.Contains(
            transaction["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "source-preserved" && check["ok"]!.GetValue<bool>());
    }

    [Fact]
    public void Capabilities_SafeInPlaceReadyRequiresRhwpAndApiRuntime()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", CreateFakeRhwp());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", null);

        var (_, missingApiStdout) = InvokeOfficeCli(["capabilities", "--json"]);
        var missingApi = JsonNode.Parse(missingApiStdout)!["data"]!["formats"]!["hwp"]!["operations"]!["replace_text"]!["safeInPlace"]!;
        Assert.False(missingApi["ready"]!.GetValue<bool>());

        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var (_, readyStdout) = InvokeOfficeCli(["capabilities", "--json"]);
        var ready = JsonNode.Parse(readyStdout)!["data"]!["formats"]!["hwp"]!["operations"]!["replace_text"]!["safeInPlace"]!;
        Assert.True(ready["ready"]!.GetValue<bool>());
    }

    private static string Sha256(string path)
    {
        var hash = SHA256.HashData(File.ReadAllBytes(path));
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}
