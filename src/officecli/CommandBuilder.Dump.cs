// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildDumpCommand(Option<bool> jsonOption)
    {
        var dumpFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx or .pptx)" };
        var dumpPathArg = new Argument<string>("path")
        {
            Description = "DOM path of the subtree to dump. Defaults to '/' (whole document) when omitted. "
                        + "Supported docx subtree paths: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles. "
                        + "Supported pptx subtree paths: /, /presentation, /slide[N], /theme, /notesMaster, /slideMaster[N], /slideLayout[N], /noteSlide[N]. "
                        + "Subtree dumps do NOT include resources at sibling paths (styles/numbering/theme; pptx: master/layout/theme); replay target must already define referenced styles/numIds/layouts.",
            DefaultValueFactory = _ => "/"
        };
        var formatOpt = new Option<string>("--format")
        {
            Description = "Output format (currently: batch)",
            DefaultValueFactory = _ => "batch"
        };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Write output to a file instead of stdout" };

        var dumpCommand = new Command("dump", "Serialize a document subtree into a replayable batch script (round-trip mechanism)");
        dumpCommand.Add(dumpFileArg);
        dumpCommand.Add(dumpPathArg);
        dumpCommand.Add(formatOpt);
        dumpCommand.Add(outOpt);
        dumpCommand.Add(jsonOption);

        dumpCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(dumpFileArg)!;
            var path = result.GetValue(dumpPathArg) ?? "/";
            var format = (result.GetValue(formatOpt) ?? "batch").ToLowerInvariant();
            var outPath = result.GetValue(outOpt);

            if (format != "batch")
                throw new CliException($"Unsupported --format: {format}. Valid: batch")
                    { Code = "invalid_format", ValidValues = ["batch"] };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".docx" && ext != ".pptx")
                throw new CliException($"dump currently supports .docx and .pptx (got {ext})")
                    { Code = "unsupported_format" };

            // CONSISTENCY(file-not-found): mirror the get/set/query format —
            // "File not found: <path>. Use 'officecli new <path>' to create a
            // blank document, or check the file extension.". Without this
            // early guard the dump path falls through to the SDK opener whose
            // raw '.NET Could not find file' message disagrees with every
            // other command and skips the actionable suggestion.
            if (!File.Exists(file.FullName))
                throw new CliException(
                    $"File not found: {file.FullName}. " +
                    $"Use 'officecli new {file.FullName}' to create a blank document, " +
                    $"or check the file extension.")
                    { Code = "file_not_found" };

            // BUG-DUMP-R6-01: route through the resident if one holds the file.
            // Without this, dump opens its own handler and collides with
            // the resident's lock ("file being used by another process").
            // Mirrors the TryResident calls in `get`/`query`/`set`.
            if (TryResident(file.FullName, req =>
            {
                req.Command = "dump";
                req.Json = json;
                req.Args["path"] = path;
                req.Args["format"] = format;
                if (!string.IsNullOrEmpty(outPath)) req.Args["out"] = outPath!;
            }, json) is {} rc) return rc;

            // CONSISTENCY(dump-format-dispatch): mirrors docx vs pptx branch
            List<BatchItem> items;
            List<CliWarning>? dumpWarnings = null;
            // BUG-R4-01: route open through DocumentHandlerFactory so the
            // FileFormatException / OpenXmlPackageException → CliException
            // (code=corrupt_file) wrapping applies. Without this, direct
            // `new WordHandler(...)` / `new PowerPointHandler(...)` leaks the
            // raw OOXML SDK exception out of programmatic callers (tests,
            // resident batch) — SafeRun catches it at the CLI surface but
            // any in-process consumer sees the unwrapped form.
            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
            if (ext == ".docx")
            {
                var word = (WordHandler)handler;
                items = WordBatchEmitter.EmitWord(word, path);
            }
            else // .pptx
            {
                var ppt = (PowerPointHandler)handler;
                var (pItems, pWarnings) = PptxBatchEmitter.EmitPptx(ppt, path);
                items = pItems;
                if (pWarnings.Count > 0)
                {
                    dumpWarnings = new List<CliWarning>(pWarnings.Count);
                    foreach (var w in pWarnings)
                    {
                        dumpWarnings.Add(new CliWarning
                        {
                            Message = $"skipped {w.Element} on {w.SlidePath}: {w.Reason}",
                            Code = "unsupported_element"
                        });
                        // Human-visible stderr line; resident's BuildWarnings
                        // also picks this up so resident-routed callers see
                        // an equivalent envelope.warnings entry.
                        Console.Error.WriteLine($"warning: skipped {w.Element} on {w.SlidePath}");
                    }
                }
            }

            // Compact JSON (single line) is the canonical batch wire form:
            // `batch run` consumes it directly and AI tooling pipes it through
            // jq/grep without caring about indentation. We previously
            // constructed a JsonSerializerOptions{WriteIndented=true} that was
            // never threaded into Serialize — kept the compact behavior, just
            // dropped the dead options block.
            var output = JsonSerializer.Serialize(items, BatchJsonContext.Default.ListBatchItem);
            // BUG-R4-FUZZ-3: Unix convention — `--out -` means stdout, not a
            // file literally named "-". Without this, running `dump --out -`
            // silently created a `-` file in the cwd (and could pollute the
            // project tree if invoked from inside it).
            if (outPath == "-")
                outPath = null;
            if (outPath != null)
            {
                // The on-disk file is the canonical batch wire form (bare
                // JSON array) so it can feed `batch --input <file>`
                // unchanged — wrapping it in an envelope would break
                // batch consumption.
                // CONSISTENCY(trailing-newline): stdout always ends with a
                // newline (Console.WriteLine); pair it on the file form too
                // so tools like `wc -l`, `git diff`, POSIX text-file
                // expectations and pipe-vs-file consumers agree on the
                // payload shape.
                File.WriteAllText(outPath, output + "\n");
                if (json)
                {
                    // BUG-R6-01: previously stdout returned
                    //   {"success": true, "data": "/tmp/out.json"}
                    // which was indistinguishable in shape from the
                    // no-out form (data is array). Make the file mode's
                    // envelope unambiguous by surfacing structured
                    // metadata under `data` instead of a bare path
                    // string. Callers can detect "data has outputFile" to
                    // disambiguate.
                    var meta = new System.Text.Json.Nodes.JsonObject
                    {
                        ["outputFile"] = outPath,
                        ["itemCount"] = items.Count
                    };
                    Console.WriteLine(OutputFormatter.WrapEnvelope(meta.ToJsonString(), dumpWarnings));
                }
                else
                    Console.WriteLine(outPath);
            }
            else
            {
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelope(output, dumpWarnings));
                else
                    Console.WriteLine(output);
            }
            return 0;
        }, json); });

        return dumpCommand;
    }
}
