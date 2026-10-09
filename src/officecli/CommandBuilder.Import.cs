// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildImportCommand(Option<bool> jsonOption)
    {
        var importFileArg = new Argument<FileInfo>("file") { Description = "Target Excel file (.xlsx)" };
        var importParentPathArg = new Argument<string>("parent-path") { Description = "Sheet path (e.g. /Sheet1)" };
        var importSourceArg = new Argument<FileInfo?>("source-file") { Description = "Source CSV/TSV file to import (positional, alternative to --file)" };
        importSourceArg.DefaultValueFactory = _ => null!;
        var importSourceOpt = new Option<FileInfo?>("--file") { Description = "Source CSV/TSV file to import" };
        var importStdinOpt = new Option<bool>("--stdin") { Description = "Read CSV/TSV data from stdin" };
        var importFormatOpt = new Option<string?>("--format") { Description = "Data format: csv or tsv (default: inferred from file extension, or csv)" };
        var importDelimiterOpt = new Option<string?>("--delimiter") { Description = "Field separator, one character — overrides --format and the file extension. For a CSV that is not comma-separated, e.g. the ';' files Excel exports in de-DE / ru-RU and other non-US locales. Takes a literal character (';', '|') or the escape '\\t' / 'tab'. A quote or newline is refused: the CSV reader gives those its own meaning." };
        var importDecimalOpt = new Option<string?>("--decimal") { Description = "Decimal mark the SOURCE file uses: '.' (default) or ','. Declaring ',' also makes '.' the thousands group, so \"1.234,5\" imports as 1234.5 and \"1,5\" as 1.5. Without it a decimal comma is left as text rather than guessed at — \"1,234\" is 1234 under one convention and 1.234 under the other. Usually paired with --delimiter ';', since a locale that writes 1,5 needs a non-comma separator." };
        var importHeaderOpt = new Option<bool>("--header") { Description = "First row is header: set AutoFilter, freeze pane, and repeat it on every printed page (print titles; an existing print-titles setting is kept)" };
        var importStartCellOpt = new Option<string>("--start-cell") { Description = "Starting cell (default: A1)" };
        importStartCellOpt.DefaultValueFactory = _ => "A1";

        var importCommand = new Command("import", "Import CSV/TSV data into an Excel sheet");
        importCommand.Add(importFileArg);
        importCommand.Add(importParentPathArg);
        importCommand.Add(importSourceArg);
        importCommand.Add(importSourceOpt);
        importCommand.Add(importStdinOpt);
        importCommand.Add(importFormatOpt);
        importCommand.Add(importDelimiterOpt);
        importCommand.Add(importDecimalOpt);
        importCommand.Add(importHeaderOpt);
        importCommand.Add(importStartCellOpt);
        importCommand.Add(jsonOption);

        importCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(importFileArg)!;
            var parentPath = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(importParentPathArg)!)!;
            var source = result.GetValue(importSourceOpt) ?? result.GetValue(importSourceArg);
            var useStdin = result.GetValue(importStdinOpt);
            var format = result.GetValue(importFormatOpt);
            var delimiterOpt = result.GetValue(importDelimiterOpt);
            var decimalOpt = result.GetValue(importDecimalOpt);
            var header = result.GetValue(importHeaderOpt);
            var startCell = result.GetValue(importStartCellOpt)!;

            if (!file.Exists)
                throw new CliException($"File not found: {file.FullName}")
                {
                    Code = "file_not_found",
                    Suggestion = $"Create the file first: officecli create \"{file.FullName}\""
                };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".xlsx")
                throw new CliException("Import is only supported for .xlsx files in V1")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .xlsx file"
                };

            // Read CSV content
            string csvContent;
            if (useStdin)
            {
                // StripBom for the same reason batch does it: File.ReadAllText
                // (the --file branch below) drops a UTF-8 BOM implicitly, the
                // stdin reader hands it through. Without this, `import --stdin`
                // fed a BOM'd CSV put a stray U+FEFF inside the first header
                // cell while `import --file` on the same bytes did not.
                csvContent = ReadStdInPayload("import", "Pass the data as a <source-file> path instead of --stdin.");
                if (csvContent.Length >= 4)
                    RejectBinaryImportSource(csvContent[0], csvContent[1], csvContent[2], csvContent[3], "on stdin");
            }
            else if (source != null)
            {
                if (!source.Exists)
                    throw new CliException($"Source file not found: {source.FullName}")
                    {
                        Code = "file_not_found"
                    };
                RejectBinaryImportSourceFile(source.FullName);
                csvContent = File.ReadAllText(source.FullName, Encoding.UTF8);
            }
            else
            {
                throw new CliException("Either --file or --stdin must be specified")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --file <path> to specify a CSV/TSV file, or --stdin to read from standard input"
                };
            }

            // Determine delimiter: --delimiter > the file's own `sep=X` line >
            // --format flag > file extension > default csv. The declaration
            // outranks --format because it is a statement about THIS file, but
            // an explicit --delimiter still wins over both.
            char delimiter = ',';
            if (!string.IsNullOrEmpty(delimiterOpt))
            {
                delimiter = ParseImportDelimiter(delimiterOpt);
            }
            else if (Core.CsvSepDeclaration.TryRead(csvContent, out var declaredSep, out _))
            {
                delimiter = declaredSep;
            }
            // (the declaration is stripped from the data by ExcelHandler.Import)
            else if (!string.IsNullOrEmpty(format))
            {
                delimiter = format.ToLowerInvariant() switch
                {
                    "tsv" => '\t',
                    "csv" => ',',
                    _ => throw new CliException($"Unknown format: {format}. Use 'csv' or 'tsv'")
                    {
                        Code = "invalid_value",
                        ValidValues = ["csv", "tsv"]
                    }
                };
            }
            else if (source != null)
            {
                var sourceExt = Path.GetExtension(source.FullName).ToLowerInvariant();
                if (sourceExt == ".tsv" || sourceExt == ".tab")
                    delimiter = '\t';
            }

            var decimalSeparator = ParseImportDecimal(decimalOpt, delimiter);
            // Judge the first DATA line: a `sep=X` declaration always contains
            // its own separator, so leaving it in made the warning describe the
            // declaration and then advise a delimiter the user had just chosen.
            var contentForWarning = Core.CsvSepDeclaration.TryRead(csvContent, out _, out var afterDecl)
                ? afterDecl : csvContent;
            if (LikelyWrongDelimiterWarning(contentForWarning, delimiter) is { } delimWarn)
                Console.Error.WriteLine(delimWarn);

            // Release any running resident's file lock before direct-open (import bypasses resident)
            ResidentClient.SendClose(file.FullName);
            using var handler = new OfficeCli.Handlers.ExcelHandler(file.FullName, editable: true);
            var msg = handler.Import(parentPath, csvContent, delimiter, header, startCell, decimalSeparator);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else
                Console.WriteLine(msg);
            return 0;
        }, json); });

        return importCommand;
    }

    /// <summary>
    /// One character, or the escape <c>\t</c> / <c>tab</c> for a tab. Quote and
    /// newline are refused because ExcelHandler.ParseCsv gives them structural
    /// meaning — a '"' opens a quoted field and '\r' / '\n' end a row — so a
    /// delimiter of either would fight the branch that handles it. Every other
    /// character is compared only outside quotes, so it is safe.
    /// </summary>
    internal static char ParseImportDelimiter(string raw)
    {
        var value = raw switch
        {
            "\\t" or "tab" or "TAB" or "\t" => "\t",
            _ => raw,
        };
        if (value.Length != 1)
            throw new CliException(
                $"--delimiter must be a single character, got '{raw}' ({value.Length} chars). "
                + "Use a literal separator like ';' or '|', or the escape '\\t' for a tab.")
            { Code = "invalid_value" };
        var c = value[0];
        if (c is '"' or '\n' or '\r')
            throw new CliException(
                "--delimiter cannot be a quote or a newline: the CSV reader uses those for "
                + "quoted fields and row breaks.")
            { Code = "invalid_value" };
        return c;
    }

    /// <summary>
    /// '.' (default) or ','. Refuses a decimal mark equal to the field
    /// separator, which would make every row ambiguous.
    /// </summary>
    internal static char ParseImportDecimal(string? raw, char delimiter)
    {
        if (string.IsNullOrEmpty(raw)) return '.';
        if (raw.Length != 1 || (raw[0] != '.' && raw[0] != ','))
            throw new CliException($"--decimal must be '.' or ',', got '{raw}'.")
            { Code = "invalid_value", ValidValues = [".", ","] };
        if (raw[0] == delimiter)
            throw new CliException(
                $"--decimal '{raw}' is also the field separator, so every row would be ambiguous. "
                + "A file with decimal commas needs a different separator, e.g. --delimiter ';'.")
            { Code = "invalid_value" };
        return raw[0];
    }

    /// <summary>
    /// A CSV whose separator is not the one we are about to split on imports as
    /// one fat column and reports "Imported N rows x 1 cols" — success-shaped
    /// output over wrong structure (issue #352). Say so when the first line has
    /// none of the chosen separator but does carry a common alternative. The
    /// import still runs: a genuine one-column file is legal.
    /// </summary>
    internal static string? LikelyWrongDelimiterWarning(string content, char delimiter)
    {
        var firstLine = content.Split('\n').FirstOrDefault(l => !string.IsNullOrWhiteSpace(l));
        if (firstLine == null || firstLine.Contains(delimiter)) return null;

        static string Show(char c) => c == '\t' ? "\\t" : c.ToString();
        foreach (var candidate in new[] { ';', '\t', '|' })
        {
            if (candidate == delimiter || !firstLine.Contains(candidate)) continue;
            return $"Warning: the first line contains no '{Show(delimiter)}' but does contain "
                + $"'{Show(candidate)}' — the file may be {Show(candidate)}-separated and will "
                + $"import as a single column. Pass --delimiter '{Show(candidate)}' if so.";
        }
        return null;
    }

    /// <summary>
    /// Issue #362: `import` reads CSV/TSV, but the natural mistake is to hand it a
    /// real workbook ("import this sheet from that .xlsx"). The container's bytes
    /// then went to the CSV reader and failed deep in cell validation with
    /// "cell value at A1 contains XML-illegal control character U+0003 at
    /// position 2" — the zip magic PK\x03\x04 read as text. Detect the container
    /// up-front and say what to do instead.
    /// </summary>
    private static void RejectBinaryImportSourceFile(string path)
    {
        byte[] head = new byte[4];
        int read;
        try
        {
            using var fs = File.OpenRead(path);
            read = fs.Read(head, 0, 4);
        }
        catch { return; } // unreadable — let the normal read surface the real error
        if (read < 4) return;
        RejectBinaryImportSource(head[0], head[1], head[2], head[3], $"'{Path.GetFileName(path)}'");
    }

    /// <summary>Shared magic-byte verdict for the --file and --stdin import sources.</summary>
    private static void RejectBinaryImportSource(int b0, int b1, int b2, int b3, string label)
    {
        string? kind =
            (b0 == 0x50 && b1 == 0x4B && b2 == 0x03 && b3 == 0x04)
                ? "an OOXML/zip container (.xlsx / .docx / .pptx)"
            : (b0 == 0xD0 && b1 == 0xCF && b2 == 0x11 && b3 == 0xE0)
                ? "a legacy OLE compound file (.xls / .doc / .ppt)"
            : null;
        if (kind == null) return;
        throw new CliException(
            $"Import source {label} is {kind}, not CSV/TSV text.")
        {
            Code = "unsupported_type",
            Suggestion = "import reads CSV/TSV only. Export the sheet to CSV first "
                + "(Excel: File > Save As > CSV UTF-8), then import that file. To copy content "
                + "between workbooks instead, run `dump` on the source and `batch` on the target.",
        };
    }

    private static Command BuildCreateCommand(Option<bool> jsonOption)
    {
        var createFileArg = new Argument<string>("file") { Description = "Output file path (.docx, .xlsx, .pptx)" };
        var createTypeOpt = new Option<string>("--type") { Description = "Document type (docx, xlsx, pptx) — optional, inferred from file extension" };
        var createForceOpt = new Option<bool>("--force") { Description = "Overwrite an existing file." };
        var createLocaleOpt = new Option<string>("--locale") { Description = "Locale tag (e.g. zh-CN, ja, ko, ar, he) — sets per-script default fonts in docDefaults and enables RTL layout for Arabic / Hebrew / Persian / Urdu and similar locales. Without this flag, the OS user culture (CFLocale on macOS, $LANG on Linux, user UI culture on Windows) is used as the default. Pass --locale en-US to force a deterministic LTR/Latin baseline regardless of the host machine. Default fonts are set for .docx (per-script + RTL), .xlsx and .pptx (theme East-Asian / complex-script fonts; RTL layout is docx-only)." };
        var createMinimalOpt = new Option<bool>("--minimal") { Description = "(.docx only) Skip Word's Normal.dotm-style baseline (Calibri 11pt + Normal style + theme1.xml) and emit a raw OOXML-spec docx instead. Use for testing edge cases or producing maximally compact output. Without this flag, the doc carries Word-aligned defaults so it renders identically in Word, other producers, and the cli preview." };
        var createCommand = new Command("create", "Create a blank Office document");
        createCommand.Add(createFileArg);
        createCommand.Add(createTypeOpt);
        createCommand.Add(createForceOpt);
        createCommand.Add(createLocaleOpt);
        createCommand.Add(createMinimalOpt);
        createCommand.Add(jsonOption);

        createCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(createFileArg)!;
            var type = result.GetValue(createTypeOpt);
            var force = result.GetValue(createForceOpt);
            var explicitLocale = result.GetValue(createLocaleOpt);
            var minimal = result.GetValue(createMinimalOpt);

            // Fall back to OS user culture when --locale is not explicitly
            // given. Empty / C / POSIX cultures yield null (no locale baked)
            // so CI environments produce neutral output.
            var locale = OfficeCli.Core.LocaleFontRegistry.ResolveEffectiveLocale(explicitLocale);
            bool localeInferred = !string.IsNullOrWhiteSpace(locale)
                && string.IsNullOrWhiteSpace(explicitLocale);

            // If file has no extension but --type is provided, append it
            if (!string.IsNullOrEmpty(type) && string.IsNullOrEmpty(Path.GetExtension(file)))
            {
                var ext = type.StartsWith('.') ? type : "." + type;
                file += ext;
            }

            // Check if the file is held by a resident process
            var fullPath = Path.GetFullPath(file);
            if (ResidentClient.TryConnect(fullPath, out _))
            {
                // Stale-resident recovery: if the on-disk file is gone (typical
                // example-script pattern: os.remove(FILE) then `create`), the
                // resident is pinning a path that no longer exists. Auto-close
                // it and proceed — refusing here would force every example
                // script to wrap `create` in a defensive `close`.
                if (!File.Exists(fullPath))
                {
                    ResidentClient.SendClose(fullPath);
                }
                else
                {
                    throw new CliException($"{Path.GetFileName(file)} is currently opened by a resident process. Please run 'officecli close \"{file}\"' first.")
                    {
                        Code = "file_locked",
                        Suggestion = $"Run: officecli close \"{file}\""
                    };
                }
            }

            // Refuse to silently overwrite an existing file unless --force is set.
            // OpenXML SDK's Create truncates the target otherwise, which can destroy
            // user data when an AI agent retries or mis-types the path.
            if (File.Exists(fullPath) && !force)
            {
                throw new CliException($"File already exists: {file}. Use --force to overwrite.")
                {
                    Code = "file_exists",
                    Suggestion = "Add --force flag or remove the file first."
                };
            }
            if (File.Exists(fullPath) && force)
            {
                Console.Error.WriteLine($"Overwriting existing file: {file}");
            }

            OfficeCli.BlankDocCreator.Create(file, locale, minimal);
            var fullCreatedPath = Path.GetFullPath(file);

            // If a --force overwrite replaced a file that currently has a live
            // watch session, push a full SSE refresh so the preview reflects the
            // new (blank) document instead of the stale pre-overwrite content
            // (issue #169). create replaces the whole file, so a full re-render
            // is the only correct shape — mirrors swap / refresh. Only reachable
            // when no resident pins the file (otherwise create fails file_locked
            // above); the watch server itself never opens the file. Best-effort:
            // a preview-refresh failure must never fail the create itself.
            if (WatchServer.IsWatching(fullCreatedPath))
            {
                try
                {
                    using var watchHandler = OfficeCli.Handlers.DocumentHandlerFactory.Open(fullCreatedPath, editable: false);
                    NotifyWatch(watchHandler, fullCreatedPath, null);
                }
                catch { /* preview refresh is best-effort; the file is already written */ }
            }

            // Best-effort: auto-start a short-lived resident process so
            // follow-up commands on this freshly-created file hit the
            // in-memory handler instead of re-opening from disk each time.
            // Uses a 60s idle timeout (much shorter than `open`'s default
            // 12min) so a stray `create` with no follow-up exits quickly.
            // Failure here does NOT fail the command — the file is already
            // on disk and all other commands still work via direct open.
            var noAuto = Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT");
            string? residentErr = null;
            var residentStarted = noAuto == "1" || string.Equals(noAuto, "true", StringComparison.OrdinalIgnoreCase)
                ? false
                : TryStartResidentProcess(fullCreatedPath, idleSeconds: 60, out residentErr);
            var residentSuffix = residentStarted
                ? " (kept open in background for faster subsequent commands)"
                : "";

            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelopeText($"Created: {fullCreatedPath}{residentSuffix}"));
            }
            else
            {
                Console.WriteLine($"Created: {file}{residentSuffix}");
                // Surface the inferred locale on stderr so the user can see
                // when the OS culture shaped the doc (RTL layout, CJK fonts,
                // etc.). Stays out of stdout / JSON envelope so scripts that
                // pipe `create` output aren't disturbed.
                {
                    var ext0 = Path.GetExtension(file).ToLowerInvariant();
                    bool localized = ext0 is ".docx" or ".xlsx" or ".pptx";
                    if (localeInferred && localized)
                    {
                        // RTL layout is only applied for docx; xlsx/pptx get the
                        // locale default fonts (theme EA/CS) but no RTL layout pass.
                        var rtlNote = ext0 == ".docx" && OfficeCli.Core.LocaleFontRegistry.IsRightToLeft(locale)
                            ? " (RTL layout enabled)" : "";
                        Console.Error.WriteLine($"Note: locale '{locale}' inferred from OS user culture{rtlNote}. Pass --locale to override.");
                    }
                }
                if (!residentStarted && !string.IsNullOrEmpty(residentErr))
                {
                    Console.Error.WriteLine($"Note: resident auto-start failed ({residentErr}); falling back to direct file access.");
                }
                if (Path.GetExtension(file).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  totalSlides: 0");
                    // Pair the unit so both dimensions agree (matches Get /
                    // readback after R40 — paired emit avoids mixing pt+cm).
                    var (cWStr, cHStr) = Core.EmuConverter.FormatEmuPaired(12192000, 6858000);
                    Console.WriteLine($"  slideWidth: {cWStr}");
                    Console.WriteLine($"  slideHeight: {cHStr}");
                }
            }
            return 0;
        }, json); });

        return createCommand;
    }

    private static Command BuildMergeCommand(Option<bool> jsonOption)
    {
        var mergeTemplateArg = new Argument<string>("template") { Description = "Template file path (.docx, .xlsx, .pptx) with {{key}} placeholders" };
        var mergeOutputArg = new Argument<string>("output") { Description = "Output file path" };
        var mergeDataOpt = new Option<string>("--data") { Description = "JSON data or path to .json file", Required = true };
        var mergeForceOpt = new Option<bool>("--force") { Description = "Overwrite an existing output file." };
        var mergeCommand = new Command("merge", "Merge template with JSON data, replacing {{key}} placeholders");
        mergeCommand.Add(mergeTemplateArg);
        mergeCommand.Add(mergeOutputArg);
        mergeCommand.Add(mergeDataOpt);
        mergeCommand.Add(mergeForceOpt);
        mergeCommand.Add(jsonOption);

        mergeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var template = result.GetValue(mergeTemplateArg)!;
            var output = result.GetValue(mergeOutputArg)!;
            var dataArg = result.GetValue(mergeDataOpt)!;
            var force = result.GetValue(mergeForceOpt);

            // If a resident holds the template with unsaved in-memory edits,
            // ask it to flush to disk first — otherwise File.Copy reads the
            // stale pre-edit bytes (resident saves itself; we never open the
            // file here). Mirrors import's pre-direct-open resident handshake.
            ResidentClient.SendSave(Path.GetFullPath(template));

            var data = Core.TemplateMerger.ParseMergeData(dataArg);
            var mergeResult = Core.TemplateMerger.Merge(template, output, data, force);

            if (json)
            {
                var mergeData = new System.Text.Json.Nodes.JsonObject
                {
                    ["output"] = Path.GetFullPath(output),
                    ["replacedKeys"] = mergeResult.UsedKeys.Count,
                    ["unresolvedPlaceholders"] = new System.Text.Json.Nodes.JsonArray(
                        mergeResult.UnresolvedPlaceholders.Select(p => (System.Text.Json.Nodes.JsonNode)p).ToArray())
                };
                var unresolvedCount = mergeResult.UnresolvedPlaceholders.Count;
                var message = unresolvedCount > 0
                    ? $"Merged {mergeResult.UsedKeys.Count} key(s), {unresolvedCount} unresolved placeholder(s)"
                    : $"Merged {mergeResult.UsedKeys.Count} key(s)";
                var jsonObj = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = mergeData,
                    ["message"] = message,
                };
                Console.WriteLine(jsonObj.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = false }));
            }
            else
            {
                Console.WriteLine($"Merged: {output}");
                Console.WriteLine($"  Replaced keys: {mergeResult.UsedKeys.Count}");
                if (mergeResult.UnresolvedPlaceholders.Count > 0)
                {
                    Console.Error.WriteLine($"  Warning: {mergeResult.UnresolvedPlaceholders.Count} unresolved placeholder(s):");
                    foreach (var p in mergeResult.UnresolvedPlaceholders)
                        Console.Error.WriteLine($"    - {{{{{p}}}}}");
                }
            }
            return 0;
        }, json); });

        return mergeCommand;
    }
}
