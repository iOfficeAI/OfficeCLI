using System.Diagnostics;
using OfficeCli;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
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
input=""
if [ "$#" -gt 1 ] && [ "${2#--}" = "$2" ]; then
  input="$2"
fi
out="output"
while [ "$#" -gt 0 ]; do
  if [ "$1" = "--input" ]; then
    shift
    input="$1"
  fi
  if [ "$1" = "--output" ] || [ "$1" = "-o" ]; then
    shift
    out="$1"
  fi
  shift
done
mkdir -p "$out"
if [ "$cmd" = "export-text" ]; then
  cat "$input" > "$out/page_001.txt"
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
  printf '%s %s' "$value" "$value" > "$output"
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
        File.WriteAllText(path, "before before");
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
