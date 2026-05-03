using System.Security.Cryptography;
using System.Text.Json.Nodes;
using OfficeCli;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public class HwpBridgeRealRhwpSmokeTests : IDisposable
{
    private const string ExpectedText = "가나다라마바사아ABCDEFG\n";
    private const string ExpectedSvgSha256 = "c40a1e2c759e797c6fb8032d3fe13b8ba55e25eeb395fa43636db375d6b5c01b";
    private readonly string? _oldEngine = Environment.GetEnvironmentVariable("OFFICECLI_HWP_ENGINE");
    private readonly string? _oldBridge = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH");
    private readonly string? _oldRhwp = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BIN");

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", _oldEngine);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", _oldBridge);
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", _oldRhwp);
    }

    [Fact]
    public void RealRhwpBinary_ViewTextAndSvg_MatchesSmokeGoldenWhenOptedIn()
    {
        if (OperatingSystem.IsWindows()) return;
        var realRhwp = Environment.GetEnvironmentVariable("OFFICECLI_REAL_RHWP_BIN");
        if (string.IsNullOrWhiteSpace(realRhwp) || !File.Exists(realRhwp)) return;

        var fixture = LocateRepoFile("tests/fixtures/hwp/rhwp-smoke/re-mixed-malgun-timesnew-hancom.hwp");
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BIN", realRhwp);

        var textResult = InvokeOfficeCli(["view", fixture, "text", "--json"]);
        Assert.Equal(0, textResult.ExitCode);
        var textRoot = JsonNode.Parse(textResult.Stdout)!;
        Assert.True(textRoot["success"]!.GetValue<bool>());
        Assert.Equal(ExpectedText, textRoot["data"]!["text"]!.GetValue<string>());
        Assert.Equal("rhwp-bridge", textRoot["data"]!["engine"]!.GetValue<string>());
        Assert.StartsWith("rhwp v", textRoot["data"]!["engineVersion"]!.GetValue<string>());

        var svgResult = InvokeOfficeCli(["view", fixture, "svg", "--json"]);
        Assert.Equal(0, svgResult.ExitCode);
        var svgRoot = JsonNode.Parse(svgResult.Stdout)!;
        Assert.True(svgRoot["success"]!.GetValue<bool>());
        var page = svgRoot["data"]!["pages"]![0]!;
        Assert.Equal(1, page["page"]!.GetValue<int>());
        Assert.Equal(ExpectedSvgSha256, page["sha256"]!.GetValue<string>());
        Assert.Equal(ExpectedSvgSha256, Sha256File(page["path"]!.GetValue<string>()));
    }

    private static string LocateBridgeDll()
    {
        return LocateRepoFile("src/rhwp-officecli-bridge/bin/Debug/net10.0/rhwp-officecli-bridge.dll");
    }

    private static string LocateRepoFile(string relativePath)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, relativePath);
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new FileNotFoundException($"Required test file was not found: {relativePath}");
    }

    private static (int ExitCode, string Stdout) InvokeOfficeCli(string[] args)
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

    private static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
}
