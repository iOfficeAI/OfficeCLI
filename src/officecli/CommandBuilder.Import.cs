// Copyright 2025 OfficeCli (officecli.ai)
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
        var importHeaderOpt = new Option<bool>("--header") { Description = "First row is header: set AutoFilter and freeze pane" };
        var importStartCellOpt = new Option<string>("--start-cell") { Description = "Starting cell (default: A1)" };
        importStartCellOpt.DefaultValueFactory = _ => "A1";

        var importCommand = new Command("import", "Import CSV/TSV data into an Excel sheet");
        importCommand.Add(importFileArg);
        importCommand.Add(importParentPathArg);
        importCommand.Add(importSourceArg);
        importCommand.Add(importSourceOpt);
        importCommand.Add(importStdinOpt);
        importCommand.Add(importFormatOpt);
        importCommand.Add(importHeaderOpt);
        importCommand.Add(importStartCellOpt);
        importCommand.Add(jsonOption);

        importCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(importFileArg)!;
            var parentPath = result.GetValue(importParentPathArg)!;
            var source = result.GetValue(importSourceOpt) ?? result.GetValue(importSourceArg);
            var useStdin = result.GetValue(importStdinOpt);
            var format = result.GetValue(importFormatOpt);
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
                csvContent = Console.In.ReadToEnd();
            }
            else if (source != null)
            {
                if (!source.Exists)
                    throw new CliException($"Source file not found: {source.FullName}")
                    {
                        Code = "file_not_found"
                    };
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

            // Determine delimiter: --format flag > file extension > default csv
            char delimiter = ',';
            if (!string.IsNullOrEmpty(format))
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

            // Release any running resident's file lock before direct-open (import bypasses resident)
            ResidentClient.SendClose(file.FullName);
            using var handler = new OfficeCli.Handlers.ExcelHandler(file.FullName, editable: true);
            var msg = handler.Import(parentPath, csvContent, delimiter, header, startCell);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else
                Console.WriteLine(msg);
            return 0;
        }, json); });

        return importCommand;
    }

    private static Command BuildCreateCommand(Option<bool> jsonOption)
    {
        var createFileArg = new Argument<string>("file") { Description = "Output file path (.docx, .xlsx, .pptx)" };
        var createTypeOpt = new Option<string>("--type") { Description = "Document type (docx, xlsx, pptx) — optional, inferred from file extension" };
        var createForceOpt = new Option<bool>("--force") { Description = "Overwrite an existing file." };
        var createCommand = new Command("create", "Create a blank Office document");
        createCommand.Aliases.Add("new");
        createCommand.Add(createFileArg);
        createCommand.Add(createTypeOpt);
        createCommand.Add(createForceOpt);
        createCommand.Add(jsonOption);

        createCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(createFileArg)!;
            var type = result.GetValue(createTypeOpt);
            var force = result.GetValue(createForceOpt);

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
                throw new CliException($"{Path.GetFileName(file)} is currently opened by a resident process. Please run 'officecli close \"{file}\"' first.")
                {
                    Code = "file_locked",
                    Suggestion = $"Run: officecli close \"{file}\""
                };
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

            OfficeCli.BlankDocCreator.Create(file);
            var fullCreatedPath = Path.GetFullPath(file);

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
                if (!residentStarted && !string.IsNullOrEmpty(residentErr))
                {
                    Console.Error.WriteLine($"Note: resident auto-start failed ({residentErr}); falling back to direct file access.");
                }
                if (Path.GetExtension(file).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  totalSlides: 0");
                    Console.WriteLine($"  slideWidth: {Core.EmuConverter.FormatEmu(12192000)}");
                    Console.WriteLine($"  slideHeight: {Core.EmuConverter.FormatEmu(6858000)}");
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
        var mergeCommand = new Command("merge", "Merge template with JSON data, replacing {{key}} placeholders");
        mergeCommand.Add(mergeTemplateArg);
        mergeCommand.Add(mergeOutputArg);
        mergeCommand.Add(mergeDataOpt);
        mergeCommand.Add(jsonOption);

        mergeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var template = result.GetValue(mergeTemplateArg)!;
            var output = result.GetValue(mergeOutputArg)!;
            var dataArg = result.GetValue(mergeDataOpt)!;

            var data = Core.TemplateMerger.ParseMergeData(dataArg);
            var mergeResult = Core.TemplateMerger.Merge(template, output, data);

            if (json)
            {
                var jsonObj = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["output"] = Path.GetFullPath(output),
                    ["replacedKeys"] = mergeResult.UsedKeys.Count,
                    ["unresolvedPlaceholders"] = new System.Text.Json.Nodes.JsonArray(
                        mergeResult.UnresolvedPlaceholders.Select(p => (System.Text.Json.Nodes.JsonNode)p).ToArray())
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
