using System.Diagnostics;
using System.Text.Json.Nodes;
using OfficeCli;

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

    [Fact]
    public void SetCellText_DelegatesToRhwpApiBridgeAndCreatesHwpOutput()
    {
        if (OperatingSystem.IsWindows()) return;
        var bridgeDll = LocateBridgeDll();
        var fakeApi = CreateFakeRhwpApi();
        var input = CreateInput(".hwpx");
        var output = CreateOutput(".hwp");

        var result = RunBridge(
            bridgeDll,
            [
                "set-cell-text", "--format", "hwpx", "--input", input, "--output", output,
                "--output-format", "hwp", "--section", "0", "--parent-para", "1",
                "--control", "0", "--cell", "0", "--cell-para", "0",
                "--value", "리지시점", "--json"
            ],
            fakeApi);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(result.Stdout)!;
        Assert.Equal("hwp", root["outputFormat"]!.GetValue<string>());
        Assert.Equal("보도시점", root["cell"]!["oldText"]!.GetValue<string>());
        Assert.Equal("리지시점", root["cell"]!["newText"]!.GetValue<string>());
    }

    [Fact]
    public void OfficeCliSetTableCell_RoutesBinaryHwpThroughRhwpBridge()
    {
        if (OperatingSystem.IsWindows()) return;
        Environment.SetEnvironmentVariable("OFFICECLI_HWP_ENGINE", "rhwp-experimental");
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH", LocateBridgeDll());
        Environment.SetEnvironmentVariable("OFFICECLI_RHWP_API_BIN", CreateFakeRhwpApi());
        var input = CreateInput(".hwp");
        var output = CreateOutput(".hwp");

        var (exitCode, stdout) = InvokeOfficeCli(
            [
                "set", input, "/table/cell",
                "--prop", "section=0",
                "--prop", "parent-para=1",
                "--prop", "control=0",
                "--prop", "cell=0",
                "--prop", "value=리지시점",
                "--prop", $"output={output}",
                "--json"
            ]);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(output));
        var root = JsonNode.Parse(stdout)!;
        Assert.True(root["success"]!.GetValue<bool>());
        Assert.Equal(output, root["data"]!["outputPath"]!.GetValue<string>());
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

    private string CreateFakeRhwpApi()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fake-rhwp-api-scan-{Guid.NewGuid():N}");
        File.WriteAllText(path, """
#!/bin/sh
if [ "$1" = "scan-cells" ]; then
  printf '%s\n' '{"cells":[{"section":0,"parentPara":1,"control":0,"cell":0,"cellPara":0,"text":"보도시점"}],"count":1,"limits":{"section":0,"maxParentPara":3,"maxControl":1,"maxCell":4,"maxCellPara":0},"engineVersion":"rhwp-api v0.test","format":"hwpx","warnings":["bounded scan"]}'
  exit 0
fi
if [ "$1" = "set-cell-text" ]; then
  output=""
  value=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    elif [ "$1" = "--value" ]; then
      shift
      value="$1"
    fi
    shift
  done
  printf 'fake hwp cell output' > "$output"
  printf '{"cell":{"oldText":"보도시점","newText":"%s"},"output":"%s","outputFormat":"hwp","engineVersion":"rhwp-api v0.test","format":"hwpx","warnings":["experimental set-cell-text"]}\n' "$value" "$output"
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

    private string CreateOutput(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-scan-output-{Guid.NewGuid():N}{extension}");
        _tempPaths.Add(path);
        return path;
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
