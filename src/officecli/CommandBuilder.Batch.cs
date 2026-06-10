// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildBatchCommand(Option<bool> jsonOption)
    {
        var batchFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var batchInputOpt = new Option<FileInfo?>("--input") { Description = "JSON file containing batch commands. If omitted, reads from stdin" };
        var batchCommandsOpt = new Option<string?>("--commands") { Description = "Inline JSON array of batch commands (alternative to --input or stdin)" };
        // BUG-R4-BT2: default flipped to continue-on-error. A 700-command
        // dump replay losing 80% of the document on the first failing item
        // (e.g. one unsupported prop) is a far worse default than reporting
        // the failure and letting the rest of the batch through. Errors are
        // still surfaced individually (BatchResult.Error) and the overall
        // exit code is 1 if any item failed, so callers can still tell
        // "everything succeeded". `--stop-on-error` opts back into the
        // strict abort-on-first-failure flow for callers who depend on it.
        var batchForceOpt = new Option<bool>("--force") { Description = "Deprecated alias for the default continue-on-error mode (kept for compatibility)" };
        var batchStopOpt = new Option<bool>("--stop-on-error") { Description = "Abort the batch as soon as any command fails (default: continue and report per-item errors)" };
        var batchCommand = new Command("batch", "Execute multiple commands from a JSON array (one open/save cycle)");
        batchCommand.Add(batchFileArg);
        batchCommand.Add(batchInputOpt);
        batchCommand.Add(batchCommandsOpt);
        batchCommand.Add(batchForceOpt);
        batchCommand.Add(batchStopOpt);
        batchCommand.Add(jsonOption);

        batchCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(batchFileArg)!;
            var inputFile = result.GetValue(batchInputOpt);
            var inlineCommands = result.GetValue(batchCommandsOpt);
            // Default: continue on error. --stop-on-error flips it to strict.
            // --force still acts as the docx-protection bypass (matches set
            // --force semantics) but no longer doubles as the continue-on-
            // error switch.
            var stopOnError = result.GetValue(batchStopOpt);
            var forceFlag = result.GetValue(batchForceOpt);

            string jsonText;
            // BUG-R7-09 (F-6): previously --commands/--input/stdin were
            // silently prioritized in that order — passing two of them at
            // once dropped the lower-priority source with no warning, so
            // scripts could fail subtly when an agent piped data into a
            // command that already had --commands set. Reject the
            // combination loudly. (Detect stdin via Console.IsInputRedirected
            // to avoid spurious failures from interactive terminals.)
            bool stdinHasInput = Console.IsInputRedirected;
            if (inlineCommands != null && inputFile != null)
                throw new ArgumentException(
                    "batch: --commands and --input are mutually exclusive. Pick one source.");
            // '--input -' explicitly opts INTO stdin — don't emit the
            // "stdin will be ignored" warning in that case, since stdin
            // is exactly what will be read.
            var inputIsStdinAlias = inputFile != null && inputFile.Name == "-";
            if ((inlineCommands != null || (inputFile != null && !inputIsStdinAlias)) && stdinHasInput
                && Environment.GetEnvironmentVariable("OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT") == null)
            {
                Console.Error.WriteLine(
                    "Warning: batch is reading from --commands/--input but stdin is also redirected; "
                    + "stdin will be ignored. Pass only one source to silence this warning, or set "
                    + "OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT=1.");
            }
            if (inlineCommands != null)
            {
                jsonText = inlineCommands;
            }
            else if (inputFile != null)
            {
                // Accept the conventional Unix '-' alias for stdin so
                // pipelines like `cat ops.json | officecli batch foo.pptx --input -`
                // don't have to drop --input entirely. Matches the implicit
                // "no --input ⇒ read stdin" branch below; using --input -
                // makes the intent explicit instead of relying on the
                // absent-flag default. (TargetMode/Exists checks are
                // skipped on purpose — '-' is not a path.)
                if (inputFile.Name == "-")
                {
                    jsonText = StripBom(Console.In.ReadToEnd());
                }
                else
                {
                    if (!inputFile.Exists)
                    {
                        throw new FileNotFoundException($"Input file not found: {inputFile.FullName}");
                    }
                    jsonText = File.ReadAllText(inputFile.FullName);
                }
            }
            else
            {
                // Read from stdin. File.ReadAllText auto-detects and strips
                // UTF-8 BOM; Console.In does not. Without an explicit strip,
                // `cat utf8bom.json | officecli batch foo.pptx` failed
                // System.Text.Json.Parse with "'﻿' is an invalid start of
                // a value" while `batch --input utf8bom.json` succeeded —
                // splitting the contract on the input source.
                jsonText = StripBom(Console.In.ReadToEnd());
            }

            // Pre-validate: check for unknown JSON fields before deserializing
            var jsonDoc = System.Text.Json.JsonDocument.Parse(jsonText);
            // CONSISTENCY(dump-batch-pipeline): `dump --json` wraps the
            // BatchItem array in an envelope object (`{"success":true,
            // "data":[…]}` via OutputFormatter.WrapEnvelope). The natural
            // pipeline `dump --json > out.json && batch --input out.json`
            // otherwise threw "Batch input must be a JSON array" because the
            // root is an object. Auto-unwrap when the envelope has a `data`
            // array so the pipeline just works without an extra `jq .data`
            // step. Bare-array inputs (dump without --json) are unaffected.
            if (jsonDoc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Object
                && jsonDoc.RootElement.TryGetProperty("data", out var envData)
                && envData.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                jsonText = envData.GetRawText();
                jsonDoc.Dispose();
                jsonDoc = System.Text.Json.JsonDocument.Parse(jsonText);
            }
            using var _jsonDocOwner = jsonDoc;
            var rootKind = jsonDoc.RootElement.ValueKind;
            if (rootKind != System.Text.Json.JsonValueKind.Array
                && rootKind != System.Text.Json.JsonValueKind.Null)
            {
                // BUG-R7-10: when the batch input is a JSON object/string/etc.
                // (not an array), Deserialize<List<BatchItem>> threw a generic
                // JsonException whose message exposed the C# generic type name
                // (`System.Collections.Generic.List`1[OfficeCli.BatchItem]`).
                // Convert it to a human-friendly error first so AI agents and
                // humans see a stable, model-agnostic diagnostic.
                throw new ArgumentException(
                    $"Batch input must be a JSON array. Got: {rootKind.ToString().ToLowerInvariant()}. "
                    + "Wrap a single item like [{\"command\":\"get\",\"path\":\"/\"}].");
            }
            if (jsonDoc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                int ri = 0;
                foreach (var elem in jsonDoc.RootElement.EnumerateArray())
                {
                    if (elem.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        var unknown = new List<string>();
                        foreach (var prop in elem.EnumerateObject())
                        {
                            if (!BatchItem.KnownFields.Contains(prop.Name))
                                unknown.Add(prop.Name);
                        }
                        if (unknown.Count > 0)
                            throw new ArgumentException($"batch item[{ri}]: unknown field(s) {string.Join(", ", unknown.Select(f => $"\"{f}\""))}. Valid fields: {string.Join(", ", BatchItem.KnownFields)}");
                    }
                    ri++;
                }
            }

            var items = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(jsonText, BatchJsonContext.Default.ListBatchItem) ?? new();
            // BUG-R40-B11: explicit null entries (e.g. `[null]`) deserialize
            // to a List<BatchItem> with a null slot and trip a NRE deeper in
            // ExecuteBatchItem. Reject up-front with a recognizable error
            // pointing at the offending index.
            for (int ni = 0; ni < items.Count; ni++)
            {
                if (items[ni] == null)
                    throw new ArgumentException(
                        $"batch item[{ni}] is null. Each entry must be a JSON object (e.g. {{\"command\":\"get\",\"path\":\"/\"}}).");
            }
            if (items.Count == 0)
            {
                // BUG-R6-07: empty command array previously short-circuited
                // before the file-existence check, so
                //   officecli batch /missing.docx --commands '[]' --json
                // returned a clean zero-result success instead of the
                // expected file_not_found. Validate the target file
                // exists first so empty-array semantics match the
                // non-empty path's diagnostics.
                if (!file.Exists)
                    throw new CliException($"File not found: {file.FullName}")
                        { Code = "file_not_found" };
                // BUG-R7-09: in --json mode an empty/null batch input
                // previously skipped the {"success":...,"data":{...}}
                // envelope used by the populated-array path, so AI agents
                // saw a missing `success` key. Apply the same envelope
                // wrap here for shape parity.
                if (json)
                {
                    using var sw = new System.IO.StringWriter();
                    PrintBatchResults(new List<BatchResult>(), json, 0, sw);
                    var inner = sw.ToString().TrimEnd('\n', '\r');
                    Console.WriteLine(OfficeCli.Core.OutputFormatter.WrapEnvelope(inner));
                }
                else
                {
                    PrintBatchResults(new List<BatchResult>(), json, 0);
                }
                return 0;
            }

            // BUG-FUZZER-R6-03: batch must honour the same .docx document
            // protection check that `set` enforces. Without this, a protected
            // doc could be silently modified via
            //   officecli batch protected.docx --commands '[{"command":"set",...}]'
            // even though the same set issued via the standalone `set` command
            // would be rejected. We piggy-back on `--force` (which already
            // means "ignore safety guards" for the continue-on-error path) so
            // agents that need to override protection use the same flag they
            // already know from `set --force`.
            // CONSISTENCY(docx-protection): if you change the protection
            // semantics, also update CommandBuilder.Set.cs at the matching
            // CheckDocxProtection call site.
            var force = forceFlag;
            // Document protection is gated ONCE per batch against the in-memory
            // DOM, not by reopening the file per item (the old per-item loop did
            // N full document opens — ~10s of I/O for a 16k batch). The check
            // runs where the live tree is available: the resident server checks
            // its in-memory _handler (see ExecuteBatch), and the non-resident
            // path checks just after it opens the handler (below). Reading the
            // live tree — not the on-disk copy — also keeps the gate correct
            // when a resident holds uncommitted in-memory protection changes
            // (resident sessions flush only on save/close).

            // If a resident process is running, send the entire batch as a
            // single "batch" command so it executes in one open/save cycle
            // inside the resident process (same semantics as non-resident mode).
            if (ResidentClient.TryConnect(file.FullName, out _))
            {
                var req = new ResidentRequest
                {
                    Command = "batch",
                    Json = json,
                    Args =
                    {
                        ["batchJson"] = jsonText,
                        ["force"] = force.ToString(),
                        ["stopOnError"] = stopOnError.ToString()
                    }
                };
                // CONSISTENCY(resident-two-step): long connectTimeoutMs so the
                // batch waits for its turn in the main-pipe queue instead of
                // silently timing out under load. Matches TryResident in
                // CommandBuilder.cs.
                var response = ResidentClient.TrySend(file.FullName, req, maxRetries: 3, connectTimeoutMs: 30000);
                if (response == null)
                {
                    Console.Error.WriteLine($"Resident for {file.Name} is running but the batch could not be delivered (main pipe busy or unresponsive). Retry, or run 'officecli close {file.Name}' and try again.");
                    return 3;
                }
                // The resident returns the formatted batch output directly
                if (!string.IsNullOrEmpty(response.Stdout))
                    Console.Write(response.Stdout);
                if (!string.IsNullOrEmpty(response.Stderr))
                    Console.Error.Write(response.Stderr);
                return response.ExitCode;
            }

            // Non-resident: open file once, execute all commands, save once.
            // Defer per-mutation Document.Save() so N commands serialize the
            // part once (at Dispose) instead of N times — eliminates an O(N²)
            // re-serialize that dominates large replays. Save-once is the
            // documented intent of this path; per-op Save was redundant given
            // the Dispose-time flush.
            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true, password: result.GetValue(PasswordOption));
            if (handler is OfficeCli.Handlers.WordHandler batchWh) batchWh.DeferSave = true;
            // Protection gate against the just-opened in-memory DOM (one check
            // for the whole batch; no second file open).
            if (!force && file.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                var protBlock = GetBatchProtectionBlock(handler, items);
                if (protBlock != null)
                {
                    if (json)
                        Console.WriteLine(OfficeCli.Core.OutputFormatter.WrapEnvelopeError(protBlock, new List<OfficeCli.Core.CliWarning>()));
                    else
                        Console.Error.WriteLine($"ERROR: {protBlock}");
                    return 1;
                }
            }
            var batchResults = new List<BatchResult>();
            for (int bi = 0; bi < items.Count; bi++)
            {
                var item = items[bi];
                try
                {
                    var output = ExecuteBatchItem(handler, item, json);
                    batchResults.Add(new BatchResult { Index = bi, Success = true, Output = output });
                }
                catch (Exception ex)
                {
                    batchResults.Add(new BatchResult { Index = bi, Success = false, Item = item, Error = ex.Message });
                    if (stopOnError) break;
                }
            }
            // BUG-R6-02: in --json mode the non-resident path emitted the raw
            // {"results":...,"summary":...} body while the resident path
            // wrapped it in {"success":..., "data":{...}} (resident server
            // calls OutputFormatter.WrapEnvelope on any JSON-shaped stdout).
            // Capture PrintBatchResults output and apply the same envelope
            // here so callers see the same shape regardless of resident state.
            // JSON Envelope contract: batch is a *judgment* command (root
            // CLAUDE.md "Judgment: any batch step failed -> outer false").
            // Outer envelope.success is true only when every step succeeded;
            // a single failed step flips outer to false even if siblings
            // succeeded. Per-step verdicts still ride on
            // `data.results[].success`. Exit code stays in lockstep with
            // envelope.success so CI gates and shells can rely on the single
            // signal. Two `success` fields appear in the JSON (outer batch
            // verdict, inner per-step) — disambiguate by JSON path.
            var batchSuccess = batchResults.Count == 0 || !batchResults.Any(r => !r.Success);
            if (json)
            {
                using var sw = new System.IO.StringWriter();
                PrintBatchResults(batchResults, json, items.Count, sw);
                var inner = sw.ToString().TrimEnd('\n', '\r');
                Console.WriteLine(OfficeCli.Core.OutputFormatter.WrapEnvelope(inner, success: batchSuccess));
            }
            else
            {
                PrintBatchResults(batchResults, json, items.Count);
            }
            if (batchResults.Any(r => r.Success))
                NotifyWatch(handler, file.FullName, null);
            return batchSuccess ? 0 : 1;
        }, json); });

        return batchCommand;
    }

    // UTF-8 BOM trim. File.ReadAllText handles this implicitly via
    // StreamReader's detect-encoding; Console.In feeds raw chars.
    private static string StripBom(string s)
        => !string.IsNullOrEmpty(s) && s[0] == '﻿' ? s.Substring(1) : s;
}
