// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    // Shown as the `batch` command's help/description. The flags alone don't
    // tell a caller (or an agent) what each JSON array ITEM looks like — the
    // single most common batch mistake is stuffing a whole CLI line into
    // "command" (e.g. {"command":"add /slide[1] --type shape --prop ..."}),
    // which fails with "Unknown command". Document the per-item shape and a
    // concrete example here so `help batch` actually teaches it.
    private const string BatchHelpDescription =
        "Execute multiple commands from a JSON array in a single pass. Standalone, this is one open/save cycle; through a live resident the items apply in memory and the disk write is deferred to save/close/idle-autosave — adaptive 2-10s after going idle, or before the batch returns under OFFICECLI_RESIDENT_FLUSH=each (officecli's own reads still see the changes immediately).\n\n"
        + "Each array item is an OBJECT whose \"command\" is the bare verb "
        + "(add/set/remove/move/swap/get/query/...); the verb's arguments are SIBLING fields, "
        + "not a CLI string inside \"command\". Common fields: \"parent\" (add target), "
        + "\"path\" (set/remove/get target), \"selector\" (query filter; \"path\" is accepted as an alias), \"type\" (element type for add), "
        + "\"props\" (a key->value map of --prop values), \"to\"/\"after\"/\"before\" (move), "
        + "\"path2\" (swap's second path).\n\n"
        + "Pass the array via --commands, or as the same JSON on stdin / --input <file>. Example:\n"
        + "[\n"
        + "  {\"command\":\"add\",\"parent\":\"/slide[1]\",\"type\":\"shape\",\"props\":{\"text\":\"Hi\",\"x\":\"1cm\",\"y\":\"2cm\"}},\n"
        + "  {\"command\":\"set\",\"path\":\"/slide[1]/shape[1]\",\"props\":{\"bold\":\"true\"}},\n"
        + "  {\"command\":\"remove\",\"path\":\"/slide[2]/shape[3]\"}\n"
        + "]\n\n"
        + "Text values follow the same rules as --prop on a single command: \\n starts a new "
        + "paragraph, \\v is a line break within one. A .docx dump taken before that split "
        + "encoded line breaks as \\n; replay one by putting "
        + "{\"command\":\"meta\",\"dumpVersion\":1} first and its \\n are read as line breaks again.\n\n"
        + "Clean-slate replay (dump->batch into a fresh target): `create` refuses to overwrite an "
        + "existing file (exit 1, code file_exists) and refuses one a live resident still holds "
        + "(exit 1, code file_locked). A script that ignores create's exit code then replays onto "
        + "the PREVIOUS run's document, so add style / add bookmark items fail with 'already exists'. "
        + "Reliable idiom: close -> rm -> create -> batch -> close. rm BEFORE create matters: with the "
        + "on-disk file gone, create auto-closes a resident still pinning that path; the leading close "
        + "just covers the handle-release window. (`create --force` overwrites an existing file but "
        + "does NOT release a resident lock — close first when a resident is up.)";

    /// <summary>
    /// Apply a batch of commands against an already-open handler. This is the
    /// single shared replay loop behind all three batch surfaces — the
    /// non-resident CLI path, the MCP server, and the resident server — so the
    /// try/catch/stop-on-error semantics can never drift between them again.
    ///
    /// Save deferral and protection gating are intentionally NOT done here:
    /// they differ by handler lifetime. The dispose-based callers (CLI
    /// non-resident, MCP) leave <c>DeferSave=true</c> and rely on the
    /// Dispose-time <c>FinalizeDeferredIds</c> flush — see
    /// <see cref="RunNonResidentBatch"/>; the long-lived resident saves and
    /// restores <c>DeferSave</c> and runs <c>ReconcileGlobalIds</c> itself.
    ///
    /// <paramref name="skipResidentOnlyCommands"/> is set by the resident, which
    /// already holds the file open: an <c>open</c>/<c>close</c> inside the batch
    /// would conflict, so they are reported as skipped instead of executed.
    /// </summary>
    internal static List<BatchResult> ApplyBatchItems(
        OfficeCli.Core.IDocumentHandler handler, List<BatchItem> items,
        bool stopOnError, bool json, bool skipResidentOnlyCommands = false,
        ICollection<string>? unrecognizedLatex = null)
    {
        var results = new List<BatchResult>();
        for (int bi = 0; bi < items.Count; bi++)
        {
            var item = items[bi];
            if (skipResidentOnlyCommands)
            {
                var cmd = (item.Command ?? "").ToLowerInvariant();
                if (cmd is "open" or "close")
                {
                    results.Add(new BatchResult { Index = bi, Success = true, Output = $"Skipped '{cmd}' (resident mode)" });
                    continue;
                }
            }
            try
            {
                var output = ExecuteBatchItem(handler, item, json);
                results.Add(new BatchResult { Index = bi, Success = true, Output = output });
            }
            catch (Exception ex)
            {
                results.Add(new BatchResult { Index = bi, Success = false, Item = item, Error = ex.Message, Code = OfficeCli.Core.OutputFormatter.InferErrorCode(ex) });
                if (stopOnError) break;
            }
            // BUG-BT2: per-item unrecognized-LaTeX diagnostics. The handler
            // resets LastUnrecognizedLatex on every Add/Set, so its tokens are
            // only valid for the item just executed — collect them now,
            // de-duplicated, so the caller can surface the same
            // unrecognized_latex_command warning (and exit 2) that single-shot
            // add/set produce. Without this the warnings were silently
            // swallowed by the batch path.
            if (unrecognizedLatex != null)
            {
                var toks = handler switch
                {
                    OfficeCli.Handlers.WordHandler wlx => wlx.LastUnrecognizedLatex,
                    OfficeCli.Handlers.PowerPointHandler plx => plx.LastUnrecognizedLatex,
                    _ => null,
                };
                if (toks is { Count: > 0 })
                    foreach (var t in toks)
                        if (!unrecognizedLatex.Contains(t)) unrecognizedLatex.Add(t);
            }
        }
        return results;
    }

    /// <summary>
    /// Run a batch against a freshly-opened, dispose-on-return handler (the
    /// non-resident CLI path and the MCP server). Sets <c>DeferSave</c> so the
    /// N commands serialize the part once at Dispose instead of N times — the
    /// O(N²) re-serialize that dominates large replays. The handler is NOT
    /// disposed here; the caller's <c>using</c> performs the single
    /// <c>FinalizeDeferredIds + Save</c> flush. Output formatting and the
    /// protection gate stay with the caller (their surfaces differ).
    /// </summary>
    internal static List<BatchResult> RunNonResidentBatch(
        OfficeCli.Core.IDocumentHandler handler, List<BatchItem> items,
        bool stopOnError, bool json, ICollection<string>? unrecognizedLatex = null)
    {
        if (handler is OfficeCli.Handlers.WordHandler wh) wh.DeferSave = true;
        return ApplyBatchItems(handler, items, stopOnError, json, unrecognizedLatex: unrecognizedLatex);
    }

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
        var batchStopOpt = new Option<bool>("--stop-on-error") { Description = "Abort the batch as soon as any command fails (with --best-effort: earlier items stay applied; default atomic mode: nothing is applied either way)" };
        // Atomic-by-default: a partially-applied batch is the dirtiest possible
        // outcome for a programmatic caller (the document sits in a state the
        // caller never asked for, and the failure verdict disagrees with what
        // landed on disk). Default = all-or-nothing; --best-effort restores the
        // old apply-what-succeeds semantics for callers that want partial
        // progress (e.g. lossy replays of dumps with known-unsupported items).
        var batchBestEffortOpt = new Option<bool>("--best-effort") { Description = "Apply the items that succeed even when others fail (pre-atomic legacy semantics). Default: any failure rolls back the whole batch" };
        var batchCommand = new Command("batch", BatchHelpDescription);
        batchCommand.Add(batchFileArg);
        batchCommand.Add(batchInputOpt);
        batchCommand.Add(batchCommandsOpt);
        batchCommand.Add(batchForceOpt);
        batchCommand.Add(batchStopOpt);
        batchCommand.Add(batchBestEffortOpt);
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
            var bestEffort = result.GetValue(batchBestEffortOpt);

            string jsonText;
            // BUG-R7-09 (F-6): previously --commands/--input/stdin were
            // silently prioritized in that order — passing two of them at
            // once dropped the lower-priority source with no warning, so
            // scripts could fail subtly when an agent piped data into a
            // command that already had --commands set. Reject the
            // combination loudly. (Detect stdin via Console.IsInputRedirected
            // to avoid spurious failures from interactive terminals.)
            // IsInputRedirected alone is true for every non-interactive
            // invocation (cron, CI, `< /dev/null`, systemd), so the warning
            // below fired on effectively all scripted batch runs with only
            // one source supplied. Refine: a seekable stdin (regular file or
            // /dev/null redirect) with zero length carries no second payload
            // — skip the warning. Pipes (CanSeek=false) still warn: someone
            // is actively piping data that will be ignored.
            bool stdinHasInput = Console.IsInputRedirected;
            if (stdinHasInput)
            {
                // Peek with a short timeout: /dev/null and closed stdin hit
                // EOF instantly (no payload → no warning); a pipe carrying a
                // real second payload has data ready (warn); an open-but-idle
                // pipe times out and is treated as no payload — batch never
                // reads stdin on this path anyway, so nothing is lost. The
                // possibly-blocked Peek thread is abandoned; the process
                // exits normally.
                try
                {
                    var stdinPeek = System.Threading.Tasks.Task.Run(() =>
                    {
                        try { return StdIn.Peek() != -1; }
                        catch { return false; }
                    });
                    stdinHasInput = stdinPeek.Wait(TimeSpan.FromMilliseconds(50)) && stdinPeek.Result;
                }
                catch { /* keep IsInputRedirected verdict */ }
            }
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
                    jsonText = StripBom(StdIn.ReadToEnd());
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
                jsonText = StripBom(StdIn.ReadToEnd());
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
            // NEWLINE-SEMANTICS-V2: strip meta items; rewrite legacy (\n = soft
            // break) docx dumps to the v2 encoding before execution.
            OfficeCli.Core.BatchCompat.PrepareForReplay(items, file.FullName);
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
            // (resident sessions flush only on save/close/idle-autosave).

            // If a resident process is running, send the entire batch as a
            // single "batch" command so all items run in one pass inside the
            // resident. NOTE: unlike the standalone path, this does NOT save to
            // disk per batch — the resident applies the items in memory and
            // defers the flush to the next save/close/idle-autosave. A reader
            // that bypasses the resident must flush first (see command-open /
            // command-batch wiki "Persisting changes").
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
                        ["stopOnError"] = stopOnError.ToString(),
                        ["bestEffort"] = bestEffort.ToString()
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
            //
            // Atomic default: run the batch against a same-directory temp COPY
            // and promote it over the original only when every item succeeded.
            // The original file is never opened for write, so no SDK
            // buffering/autosave subtlety can leak a partial state onto it —
            // and a crash mid-save corrupts only the temp. --best-effort keeps
            // the legacy run-in-place semantics; all-read-only batches skip the
            // copy (nothing to protect).
            var atomic = !bestEffort && items.Any(it => !ReadOnlyBatchVerbs.Contains(it.Command ?? ""));
            // Resolve a symlink up front so the final promote replaces the
            // TARGET file; a plain rename would overwrite the link itself.
            string targetPath;
            try { targetPath = new FileInfo(file.FullName).ResolveLinkTarget(returnFinalTarget: true)?.FullName ?? file.FullName; }
            catch { targetPath = file.FullName; }
            string? tmpPath = null;
            var workPath = targetPath;
            if (atomic)
            {
                var tmpDir = System.IO.Path.GetDirectoryName(targetPath) ?? ".";
                var tmpStem = TruncateStemForTempName(System.IO.Path.GetFileNameWithoutExtension(targetPath));
                var tmpExt = System.IO.Path.GetExtension(targetPath);
                // Sweep orphaned temp copies from earlier crashed batches
                // (kill -9 mid-run can't self-clean). Only THIS document's
                // pattern is considered, and a candidate must pass TWO gates:
                // an exclusive open (a running batch holds its temp open, so
                // live temps are never touched) AND an age floor — the
                // exclusive-open probe alone has TOCTOU windows against a
                // concurrent batch on the same document (after its File.Copy
                // but before its handler opens; after its dispose but before
                // its File.Replace) where the temp is closed yet very much
                // alive. Real crash orphans are minutes-to-days old; anything
                // younger is presumed live and left for a later sweep.
                try
                {
                    var ageFloor = DateTime.UtcNow - TimeSpan.FromMinutes(15);
                    // Two patterns: promoted temps (batch-) and staging names
                    // (batchprep-, see below) a crash may leave behind.
                    foreach (var pattern in new[] { $".{tmpStem}.batch-*{tmpExt}", $".{tmpStem}.batchprep-*{tmpExt}" })
                        foreach (var orphan in System.IO.Directory.GetFiles(tmpDir, pattern))
                        {
                            try
                            {
                                // Age gate on the NEWER of creation time and
                                // mtime. File.Copy preserves the source's
                                // mtime (possibly hours old) but the copy's
                                // birthtime is the copy moment — so a live
                                // temp/staging file is young by construction
                                // on macOS/Windows, with no stamp-ordering
                                // window to race. (On filesystems without
                                // birthtime, CreationTimeUtc falls back to a
                                // stat time and the explicit stamp below
                                // still narrows the window to copy→stamp.)
                                var bornOrWritten = System.IO.File.GetCreationTimeUtc(orphan);
                                var lastWrite = System.IO.File.GetLastWriteTimeUtc(orphan);
                                if (bornOrWritten > ageFloor || lastWrite > ageFloor) continue;
                                using (new System.IO.FileStream(orphan, System.IO.FileMode.Open,
                                    System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)) { }
                                System.IO.File.Delete(orphan);
                            }
                            catch (Exception delEx)
                            {
                                // In use (concurrent batch) is normal; an
                                // UNDELETABLE aged orphan (e.g. macOS uchg
                                // flag cloned from a locked document) would
                                // otherwise accumulate one hidden file per
                                // retry with no signal — say so once.
                                if (delEx is not System.IO.IOException)
                                    Console.Error.WriteLine($"warning: could not remove stale batch temp '{orphan}': {delEx.Message}");
                            }
                        }
                }
                catch { /* directory unreadable — sweep is best-effort */ }
                // Keep the original extension LAST — DocumentHandlerFactory
                // dispatches by extension, so `.tmp` would be rejected.
                //
                // STREAM-copy under a staging name, rename into place. Not
                // File.Copy: that preserves the source's mtime (hours old on
                // an idle document), and on macOS setting an mtime below the
                // birthtime drags the birthtime down with it — no timestamp
                // gate survives that. A streamed copy is born with a fresh
                // mtime, holds an exclusive handle for the whole write (the
                // sweep's open-probe bounces off), and inherits neither BSD
                // file flags (uchg) nor the source's timestamps. From handle
                // close to rename the file is protected by its fresh mtime.
                var guid = Guid.NewGuid().ToString("N");
                var prepPath = System.IO.Path.Combine(tmpDir, $".{tmpStem}.batchprep-{guid}{tmpExt}");
                tmpPath = System.IO.Path.Combine(tmpDir, $".{tmpStem}.batch-{guid}{tmpExt}");
                using (var src = new System.IO.FileStream(targetPath, System.IO.FileMode.Open,
                    System.IO.FileAccess.Read, System.IO.FileShare.Read))
                using (var dst = new System.IO.FileStream(prepPath, System.IO.FileMode.CreateNew,
                    System.IO.FileAccess.ReadWrite, System.IO.FileShare.None))
                {
                    src.CopyTo(dst);
                }
                // The final promote is a rename, which would hand the document
                // the temp's default permissions — carry the original mode
                // over. (Windows File.Replace preserves the destination's ACL
                // by itself.)
                if (!OperatingSystem.IsWindows())
                    try { System.IO.File.SetUnixFileMode(prepPath, System.IO.File.GetUnixFileMode(targetPath)); }
                    catch { /* best-effort */ }
                System.IO.File.Move(prepPath, tmpPath);
                workPath = tmpPath;
            }
            List<BatchResult> batchResults;
            var batchUnrecognizedLatex = new List<string>();
            bool batchSuccessLocal;
            try
            {
                using var handler = DocumentHandlerFactory.Open(workPath, editable: true);
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
                        if (tmpPath != null)
                            try { System.IO.File.Delete(tmpPath); } catch { /* best-effort */ }
                        return 1;
                    }
                }
                // DeferSave + replay loop, shared with the MCP batch surface. The
                // handler's using-Dispose performs the single FinalizeDeferredIds +
                // Save flush.
                batchResults = RunNonResidentBatch(handler, items, stopOnError, json, batchUnrecognizedLatex);
                batchSuccessLocal = batchResults.Count == 0 || !batchResults.Any(r => !r.Success);
                // Watch preview: only when the document actually changes — the
                // atomic rollback leaves the file exactly as the preview
                // already shows it, so a refresh would be a lie (and a waste).
                if (batchResults.Any(r => r.Success) && (batchSuccessLocal || !atomic))
                    NotifyWatch(handler, file.FullName, null);
            }
            catch
            {
                // Exception escaping the replay (mapped to an error envelope by
                // SafeRun): never leave the temp copy behind, never promote it.
                if (tmpPath != null)
                    try { System.IO.File.Delete(tmpPath); } catch { /* best-effort */ }
                throw;
            }
            // The using-Dispose above fully serialized the (temp) document;
            // promote it over the original only on an all-green batch.
            if (tmpPath != null)
            {
                if (batchSuccessLocal)
                {
                    // Same-volume atomic swap; the temp copy carries the fully
                    // saved post-batch document. A failing swap (target made
                    // read-only / deleted underneath us / fs quirk) surfaces as
                    // a command error BEFORE any success output — the original
                    // is untouched, so the verdict stays truthful — but the
                    // temp copy must not leak on that path.
                    try
                    {
                        System.IO.File.Replace(tmpPath, targetPath, destinationBackupFileName: null);
                    }
                    catch
                    {
                        try { System.IO.File.Delete(tmpPath); } catch { /* best-effort */ }
                        throw;
                    }
                }
                else
                {
                    try { System.IO.File.Delete(tmpPath); } catch { /* best-effort */ }
                }
            }
            // BUG-R6-02: in --json mode the non-resident path emitted the raw
            // {"results":...,"summary":...} body while the resident path
            // wrapped it in {"success":..., "data":{...}} (resident server
            // calls OutputFormatter.WrapEnvelope on any JSON-shaped stdout).
            // Capture PrintBatchResults output and apply the same envelope
            // here so callers see the same shape regardless of resident state.
            // JSON Envelope contract: batch is a *judgment* command (root
            // the project conventions "Judgment: any batch step failed -> outer false").
            // Outer envelope.success is true only when every step succeeded;
            // a single failed step flips outer to false even if siblings
            // succeeded. Per-step verdicts still ride on
            // `data.results[].success`. Exit code stays in lockstep with
            // envelope.success so CI gates and shells can rely on the single
            // signal. Two `success` fields appear in the JSON (outer batch
            // verdict, inner per-step) — disambiguate by JSON path.
            var batchSuccess = batchResults.Count == 0 || !batchResults.Any(r => !r.Success);
            // BUG-BT2: surface per-item unrecognized-LaTeX warnings the same way
            // single-shot add/set do (warning + JSON envelope + exit 2). A batch
            // whose only issue is an unknown LaTeX command otherwise exited 0
            // with no diagnostic, silently writing the literal-text fallback.
            var batchWarnings = new List<OfficeCli.Core.CliWarning>();
            foreach (var tok in batchUnrecognizedLatex)
            {
                batchWarnings.Add(new OfficeCli.Core.CliWarning
                {
                    Message = $"unrecognized_latex_command: {tok}",
                    Code = "unrecognized_latex_command",
                    Suggestion = "Check the command spelling; see https://katex.org/docs/supported.html for supported syntax.",
                });
            }
            var rolledBack = atomic && !batchSuccess;
            if (json)
            {
                using var sw = new System.IO.StringWriter();
                PrintBatchResults(batchResults, json, items.Count, sw, atomicRolledBack: rolledBack);
                var inner = sw.ToString().TrimEnd('\n', '\r');
                Console.WriteLine(OfficeCli.Core.OutputFormatter.WrapEnvelope(inner, batchWarnings, success: batchSuccess));
            }
            else
            {
                PrintBatchResults(batchResults, json, items.Count, atomicRolledBack: rolledBack);
                foreach (var w in batchWarnings)
                    Console.Error.WriteLine($"  WARNING: {w.Message}");
            }
            // (watch notify happened inside the handler scope above)
            // Exit precedence: a failed item (exit 1) outranks an
            // unrecognized-LaTeX-only warning (exit 2 mirrors single-shot).
            if (!batchSuccess) return 1;
            return batchWarnings.Count > 0 ? 2 : 0;
        }, json); });

        return batchCommand;
    }

    // UTF-8 BOM trim. File.ReadAllText handles this implicitly via
    // StreamReader's detect-encoding; the stdin reader feeds raw chars.
    private static string StripBom(string s)
        => !string.IsNullOrEmpty(s) && s[0] == '﻿' ? s.Substring(1) : s;

    /// <summary>
    /// UTF-8 view of the process's stdin, used everywhere this CLI reads piped
    /// input (batch JSON, import CSV/TSV).
    ///
    /// Console.In would decode with Console.InputEncoding, which on Windows is
    /// the console's input code page (CP437 on en-US, CP936 on zh-CN) and never
    /// UTF-8 — piped payloads arrived mojibaked. Setting Console.InputEncoding
    /// fixes the decode but calls SetConsoleCP on the console object, which is
    /// shared with the parent shell and outlives this process: every officecli
    /// run would leave the user's terminal pinned to CP 65001, breaking
    /// non-ASCII keyboard input for legacy console programs run afterwards.
    /// Gating that mutation on Console.IsInputRedirected does not help — a piped
    /// run has a console attached too, so it leaks in exactly the case the decode
    /// fix is for. Reading through our own StreamReader gets correct UTF-8 with no
    /// global console mutation. Same approach McpServer already uses.
    ///
    /// The reader itself IS gated on IsInputRedirected, for the opposite reason:
    /// with no redirect the stream sits on the console, which hands back bytes in
    /// its own input code page, so a UTF-8 reader would mojibake anything
    /// non-ASCII typed at the prompt. Console.In decodes that case correctly and
    /// there is nothing to fix there — the bug is redirected input only.
    ///
    /// One shared instance, because the batch path Peeks before it Reads and a
    /// second reader would swallow whatever the first one buffered. Wrapped in
    /// TextReader.Synchronized to match Console.In's thread-safety — the Peek
    /// runs on a worker thread that may be abandoned on timeout.
    /// </summary>
    private static readonly Lazy<TextReader> LazyStdIn = new(() =>
        Console.IsInputRedirected
            ? TextReader.Synchronized(new StreamReader(
                Console.OpenStandardInput(),
                new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                // StripBom owns BOM handling; don't let detection also switch encodings.
                detectEncodingFromByteOrderMarks: false))
            : Console.In);

    internal static TextReader StdIn => LazyStdIn.Value;

    /// <summary>
    /// The atomic temp name adds ~45 bytes of affixes around the document's
    /// stem; a stem near the 255-byte filename limit would push the temp name
    /// over it, making `batch` fail on a file that plain `set` handles.
    /// Truncate the stem (UTF-8-byte budget, at a char boundary) for the temp
    /// name AND the orphan-sweep pattern — both must derive identically or a
    /// crashed batch's orphan would never match its own document's sweep.
    /// Collision risk from truncation is absorbed by the per-temp GUID.
    /// </summary>
    private static string TruncateStemForTempName(string stem)
    {
        const int maxStemBytes = 180;
        if (System.Text.Encoding.UTF8.GetByteCount(stem) <= maxStemBytes) return stem;
        var bytes = 0;
        var end = 0;
        while (end < stem.Length)
        {
            // Walk whole code points so a surrogate pair is never split
            // (a lone surrogate cannot round-trip through a UTF-8 filename).
            var step = char.IsHighSurrogate(stem[end]) && end + 1 < stem.Length ? 2 : 1;
            var b = System.Text.Encoding.UTF8.GetByteCount(stem.AsSpan(end, step));
            if (bytes + b > maxStemBytes) break;
            bytes += b;
            end += step;
        }
        return stem[..end];
    }

    // Batch verbs that never mutate the DOM. Single source of truth for both
    // the non-resident atomic-copy decision above and the resident server's
    // watch-notify / atomic decisions (ResidentServer.ExecuteBatch). Any verb
    // NOT listed here (including unknown / future ones) fails open to
    // "mutating", so a new verb is never silently treated as read-only.
    internal static readonly HashSet<string> ReadOnlyBatchVerbs = new(StringComparer.OrdinalIgnoreCase)
    {
        "get", "query", "view", "validate", "dump", "raw"
    };
}
