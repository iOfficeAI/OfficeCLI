using System.Diagnostics;
using System.Text.Json.Nodes;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public class HwpBridgeTableScanTests : IDisposable
{
    private readonly List<string> _tempPaths = new();

    public void Dispose()
    {
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
    public void ScanCells_DelegatesToRhwpApiBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwpx");

        var result = RunBridge(
            bridgeDll,
            [
                "scan-cells", "--format", "hwpx", "--input", input,
                "--section", "0", "--max-parent-para", "3", "--max-control", "1",
                "--max-cell", "4", "--max-cell-para", "0", "--json"
            ],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal(1, root["count"]!.GetValue<int>());
        Assert.Equal("보도시점", root["cells"]![0]!["text"]!.GetValue<string>());
    }

    private static string LocateBridgeDll()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(
                dir.FullName,
                "src/rhwp-officecli-bridge/bin/Debug/net10.0/rhwp-officecli-bridge.dll");
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new FileNotFoundException("rhwp-officecli-bridge.dll was not built.");
    }

    private ProcessResult RunBridge(string bridgeDll, string[] args, string fakeRhwpApi)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        psi.ArgumentList.Add(bridgeDll);
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_RHWP_BIN"] = "/unused/rhwp";
        psi.Environment["OFFICECLI_RHWP_API_BIN"] = fakeRhwpApi;
        using var process = Process.Start(psi)!;
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    private string CreateFakeRhwpApi()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fake-rhwp-api-scan-{Guid.NewGuid():N}");
        File.WriteAllText(path, """
#!/bin/sh
if [ "$1" = "scan-cells" ]; then
  printf '%s\n' '{"cells":[{"section":0,"parentPara":1,"control":0,"cell":0,"cellPara":0,"text":"보도시점"}],"count":1,"limits":{"section":0,"maxParentPara":3,"maxControl":1,"maxCell":4,"maxCellPara":0},"engineVersion":"rhwp-api v0.test","format":"hwpx","warnings":["bounded scan"]}'
  exit 0
fi
echo "unexpected api command: $1" >&2
exit 2
""");
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        _tempPaths.Add(path);
        return path;
    }

    private string CreateInput(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-scan-input-{Guid.NewGuid():N}{extension}");
        File.WriteAllText(path, "fake");
        _tempPaths.Add(path);
        return path;
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
