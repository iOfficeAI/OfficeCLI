using System.Diagnostics;
using System.Text.Json.Nodes;
using OfficeCli;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public class HwpBridgeSidecarTests : IDisposable
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
        Assert.Equal("Page one\nPage two\n", root["text"]!.GetValue<string>());
        Assert.Equal("rhwp v0.test", root["engineVersion"]!.GetValue<string>());
        Assert.Equal("hwp", root["format"]!.GetValue<string>());
        Assert.Equal(1, root["pages"]![0]!["page"]!.GetValue<int>());
        Assert.Equal("Page one\n", root["pages"]![0]!["text"]!.GetValue<string>());
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
        Assert.Equal("Page one\nPage two\n", root["data"]!["text"]!.GetValue<string>());
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
        Assert.Equal("fake", File.ReadAllText(input));
        var root = JsonNode.Parse(stdout)!;
        Assert.False(root["success"]!.GetValue<bool>());
        Assert.Equal("in-place", root["data"]!["transaction"]!["mode"]!.GetValue<string>());
        Assert.False(root["data"]!["transaction"]!["ok"]!.GetValue<bool>());
        Assert.Contains(
            root["data"]!["transaction"]!["checks"]!.AsArray(),
            check => check?["name"]?.GetValue<string>() == "in-place-not-ready");
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

    private ProcessResult RunBridge(string bridgeDll, string fakeRhwp, string[] args, string? fakeRhwpApi = null)
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
        psi.Environment["OFFICECLI_RHWP_BIN"] = fakeRhwp;
        if (fakeRhwpApi != null) psi.Environment["OFFICECLI_RHWP_API_BIN"] = fakeRhwpApi;
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

    private string CreateFakeRhwp()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fake-rhwp-{Guid.NewGuid():N}");
        File.WriteAllText(path, """
#!/bin/sh
if [ "$1" = "--version" ]; then
  echo "rhwp v0.test"
  exit 0
fi
cmd="$1"
out="output"
while [ "$#" -gt 0 ]; do
  if [ "$1" = "--output" ] || [ "$1" = "-o" ]; then
    shift
    out="$1"
  fi
  shift
done
mkdir -p "$out"
if [ "$cmd" = "export-text" ]; then
  printf 'Page one\n' > "$out/page_001.txt"
  printf 'Page two\n' > "$out/page_002.txt"
  exit 0
fi
if [ "$cmd" = "export-svg" ]; then
  printf '<svg><text>page</text></svg>' > "$out/page.svg"
  exit 0
fi
echo "unexpected command: $cmd" >&2
exit 2
""");
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        _tempPaths.Add(path);
        return path;
    }

    private string CreateFakeRhwpApi()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fake-rhwp-api-{Guid.NewGuid():N}");
        File.WriteAllText(path, """
#!/bin/sh
cmd="$1"
if [ "$cmd" = "list-fields" ]; then
  printf '%s\n' '{"fields":[{"fieldId":7,"fieldType":"clickhere","name":"applicant","value":"홍길동"}],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "get-field" ]; then
  printf '%s\n' '{"field":{"ok":true,"fieldId":7,"value":"홍길동"},"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "set-field" ]; then
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
  printf 'fake hwp output' > "$output"
  printf '{"field":{"ok":true,"fieldId":7,"oldValue":"","newValue":"%s"},"output":"%s","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental set-field"]}\n' "$value" "$output"
  exit 0
fi
if [ "$cmd" = "replace-text" ]; then
  output=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    fi
    shift
  done
  printf 'fake hwp replace output' > "$output"
  printf '{"replacement":{"ok":true,"count":2},"output":"%s","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental replace-text"]}\n' "$output"
  exit 0
fi
if [ "$cmd" = "get-cell-text" ]; then
  printf '%s\n' '{"cell":{"section":0,"parentPara":2,"control":0,"cell":1,"cellPara":0,"offset":0,"count":1000,"text":"셀값"},"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
echo "unexpected api command: $cmd" >&2
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
        var path = Path.Combine(Path.GetTempPath(), $"bridge-input-{Guid.NewGuid():N}{extension}");
        File.WriteAllText(path, "fake");
        _tempPaths.Add(path);
        return path;
    }

    private string CreateOutput(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-output-{Guid.NewGuid():N}{extension}");
        _tempPaths.Add(path);
        return path;
    }

    private string CreateDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-out-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        _tempPaths.Add(path);
        return path;
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
