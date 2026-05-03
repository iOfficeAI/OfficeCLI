using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;

var exitCode = BridgeProgram.Run(args);
return exitCode;

internal static class BridgeProgram
{
    public static int Run(string[] args)
    {
        if (args.Length == 0 || args[0] is "--help" or "-h")
            return Help();

        try
        {
            return args[0] switch
            {
                "read-text" => ReadText(args[1..]),
                "render-svg" => RenderSvg(args[1..]),
                "list-fields" => ApiBridge(args),
                "get-field" => ApiBridge(args),
                "set-field" => ApiBridge(args),
                "replace-text" => ApiBridge(args),
                "get-cell-text" => ApiBridge(args),
                "scan-cells" => ApiBridge(args),
                "set-cell-text" => ApiBridge(args),
                _ => Error($"unsupported command: {args[0]}", "unsupported_command")
            };
        }
        catch (Exception ex)
        {
            return Error(ex.Message, "bridge_exception");
        }
    }

    private static int ReadText(string[] args)
    {
        var options = ParseOptions(args);
        var input = Required(options, "--input");
        var format = Required(options, "--format");
        if (!File.Exists(input)) return Error($"input not found: {input}", "input_not_found");

        using var temp = TempDirectory.Create();
        var rhwp = RhwpBinary();
        var rhwpArgs = new List<string> { "export-text", input, "--output", temp.Path };
        var result = RunProcess(rhwp, rhwpArgs);
        if (result.ExitCode != 0)
            return Error($"rhwp export-text failed: {result.Stderr.Trim()}", "rhwp_failed");

        var textFiles = Directory.GetFiles(temp.Path, "*.txt").OrderBy(p => p, StringComparer.Ordinal).ToArray();
        var pages = textFiles.Select((path, index) => new TextPage(index + 1, File.ReadAllText(path))).ToArray();
        var text = string.Concat(pages.Select(p => p.Text));
        WriteJson(new TextResponse(text, pages, EngineVersion(rhwp), [], format));
        return 0;
    }

    private static int RenderSvg(string[] args)
    {
        var options = ParseOptions(args);
        var input = Required(options, "--input");
        var format = Required(options, "--format");
        var outDir = Required(options, "--out-dir");
        var pageSelector = options.GetValueOrDefault("--page", "all");
        if (!File.Exists(input)) return Error($"input not found: {input}", "input_not_found");
        Directory.CreateDirectory(outDir);

        var rhwp = RhwpBinary();
        var rhwpArgs = new List<string> { "export-svg", input, "--output", outDir };
        if (!string.Equals(pageSelector, "all", StringComparison.OrdinalIgnoreCase))
        {
            if (!int.TryParse(pageSelector, out var oneBasedPage) || oneBasedPage <= 0)
                return Error($"unsupported page selector: {pageSelector}", "unsupported_page_selector");
            rhwpArgs.Add("--page");
            rhwpArgs.Add((oneBasedPage - 1).ToString());
        }

        var result = RunProcess(rhwp, rhwpArgs);
        if (result.ExitCode != 0)
            return Error($"rhwp export-svg failed: {result.Stderr.Trim()}", "rhwp_failed");

        var svgFiles = Directory.GetFiles(outDir, "*.svg").OrderBy(p => p, StringComparer.Ordinal).ToArray();
        var pages = svgFiles.Select((path, index) => new SvgPage(index + 1, path, Sha256(path))).ToArray();
        var manifest = Path.Combine(outDir, "manifest.json");
        File.WriteAllText(manifest, JsonSerializer.Serialize(new { pages }, JsonOptions()));
        WriteJson(new SvgResponse(pages, manifest, EngineVersion(rhwp), [], format));
        return 0;
    }

    private static Dictionary<string, string> ParseOptions(string[] args)
    {
        var options = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var i = 0; i < args.Length; i++)
        {
            if (!args[i].StartsWith("--", StringComparison.Ordinal)) continue;
            if (args[i] == "--json")
            {
                options[args[i]] = "true";
                continue;
            }
            if (i + 1 >= args.Length || args[i + 1].StartsWith("--", StringComparison.Ordinal))
                throw new ArgumentException($"missing value for {args[i]}");
            options[args[i]] = args[i + 1];
            i++;
        }
        return options;
    }

    private static string Required(Dictionary<string, string> options, string key)
        => options.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new ArgumentException($"missing required option: {key}");

    private static string RhwpBinary()
        => Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BIN") ?? "rhwp";

    private static string RhwpApiBinary()
        => Environment.GetEnvironmentVariable("OFFICECLI_RHWP_API_BIN") ?? "";

    private static int ApiBridge(string[] args)
    {
        var api = RhwpApiBinary();
        if (string.IsNullOrWhiteSpace(api) || !File.Exists(api))
            return Error(
                "rhwp API bridge is not configured. Set OFFICECLI_RHWP_API_BIN to rhwp-field-bridge.",
                "api_bridge_missing");

        var result = RunProcess(api, args);
        if (result.ExitCode != 0)
            return Error($"rhwp API bridge failed: {result.Stderr.Trim()}", "api_bridge_failed");

        Console.Write(result.Stdout);
        if (!result.Stdout.EndsWith('\n')) Console.WriteLine();
        return 0;
    }

    private static ProcessResult RunProcess(string fileName, IReadOnlyList<string> args)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        using var process = Process.Start(psi) ?? throw new InvalidOperationException($"failed to start {fileName}");
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    private static string? EngineVersion(string rhwp)
    {
        try
        {
            var result = RunProcess(rhwp, ["--version"]);
            return result.ExitCode == 0 ? result.Stdout.Trim() : null;
        }
        catch
        {
            return null;
        }
    }

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static int Error(string message, string code)
    {
        WriteJson(new ErrorResponse(false, new BridgeError(message, code)));
        Console.Error.WriteLine(message);
        return 1;
    }

    private static int Help()
    {
        Console.WriteLine("rhwp-officecli-bridge read-text|render-svg|list-fields|get-field|set-field|replace-text|get-cell-text|scan-cells|set-cell-text --format hwp|hwpx --input <path> --json");
        return 0;
    }

    private static void WriteJson<T>(T value)
        => Console.WriteLine(JsonSerializer.Serialize(value, JsonOptions()));

    private static JsonSerializerOptions JsonOptions()
        => new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
}

internal sealed class TempDirectory : IDisposable
{
    private TempDirectory(string path) => Path = path;
    public string Path { get; }
    public static TempDirectory Create()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officecli_rhwp_bridge_{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return new TempDirectory(path);
    }
    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); } catch { }
    }
}

internal sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
internal sealed record TextPage(int Page, string Text);
internal sealed record SvgPage(int Page, string Path, string Sha256);
internal sealed record TextResponse(string Text, IReadOnlyList<TextPage> Pages, string? EngineVersion, string[] Warnings, string Format);
internal sealed record SvgResponse(IReadOnlyList<SvgPage> Pages, string Manifest, string? EngineVersion, string[] Warnings, string Format);
internal sealed record ErrorResponse(bool Success, BridgeError Error);
internal sealed record BridgeError(string Message, string Code);
