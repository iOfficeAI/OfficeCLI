using System.Diagnostics;
using OfficeCli;
using OfficeCli.Tests.Hwpx;

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
  if [ "${input##*.}" = "hwpx" ]; then
    python3 - "$input" "$out/page_001.txt" <<'PY'
import html
import re
import sys
import zipfile

input_path, output_path = sys.argv[1], sys.argv[2]
texts = []
with zipfile.ZipFile(input_path) as archive:
    for name in sorted(archive.namelist()):
        if name.startswith("Contents/section") and name.endswith(".xml"):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            texts.extend(html.unescape(value) for value in re.findall(r"<(?:\w+:)?t[^>]*>(.*?)</(?:\w+:)?t>", xml, re.S))
with open(output_path, "w", encoding="utf-8") as output:
    output.write(" ".join(texts))
PY
  else
    cat "$input" > "$out/page_001.txt"
  fi
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
  format="hwp"
  input=""
  output=""
  query=""
  value=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--format" ]; then
      shift
      format="$1"
    elif [ "$1" = "--input" ]; then
      shift
      input="$1"
    elif [ "$1" = "--output" ]; then
      shift
      output="$1"
    elif [ "$1" = "--query" ]; then
      shift
      query="$1"
    elif [ "$1" = "--value" ]; then
      shift
      value="$1"
    fi
    shift
  done
  if [ "$format" = "hwpx" ]; then
    python3 - "$input" "$output" "$query" "$value" <<'PY'
import sys
import zipfile

input_path, output_path, query, value = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
with zipfile.ZipFile(input_path) as source, zipfile.ZipFile(output_path, "w") as target:
    for info in source.infolist():
        data = source.read(info.filename)
        if info.filename.endswith((".xml", ".hpf", ".rdf", ".txt")):
            text = data.decode("utf-8", errors="ignore")
            data = text.replace(query, value).encode("utf-8")
        target.writestr(info, data)
PY
  else
    printf '%s %s' "$value" "$value" > "$output"
  fi
  printf '{"replacement":{"ok":true,"count":2},"output":"%s","engineVersion":"rhwp-api v0.test","format":"%s","warnings":["experimental replace-text"]}\n' "$output" "$format"
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
