// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
// Test-only stub exporter for deck export-pdf smoke (not a real PDF renderer).

using System.Text.Json;

if (args is ["--info"])
{
    Console.WriteLine(JsonSerializer.Serialize(new
    {
        name = "officecli-pdf-stub",
        version = "0.0.1",
        protocol = 1,
        kinds = new[] { "exporter" },
        extensions = new[] { ".pdf" },
        supports = new[] { "from:pptx", "from:docx", "from:xlsx" },
        idle_timeout_seconds = new { verbs = new { export = 30 } },
        runtime = "dotnet",
    }));
    return 0;
}

if (args.Length < 1 || !string.Equals(args[0], "export", StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("usage: FakePdfExporter export <source> --out <pdf>");
    return 1;
}

string? source = null;
string? output = null;
for (var i = 1; i < args.Length; i++)
{
    if (args[i] == "--out" && i + 1 < args.Length)
    {
        output = args[++i];
        continue;
    }
    if (source is null)
        source = args[i];
}

if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(output))
{
    Console.Error.WriteLine("export requires <source> and --out <pdf>");
    return 1;
}

if (!File.Exists(source))
{
    Console.Error.WriteLine($"source missing: {source}");
    return 2;
}

var dir = Path.GetDirectoryName(Path.GetFullPath(output));
if (!string.IsNullOrEmpty(dir))
    Directory.CreateDirectory(dir);

// Minimal PDF header so callers can assert a file was written. Not a real render.
await File.WriteAllBytesAsync(output, "%PDF-1.4\n%OfficeCLI deck export-pdf stub\n"u8.ToArray());
return 0;
