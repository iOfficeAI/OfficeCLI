// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    // ==================== mark ====================

    // Canonical prop names accepted by `mark --prop`. Any other key triggers
    // the unknown-prop warning. Lower-case for case-insensitive comparison
    // (the prop dictionary itself is OrdinalIgnoreCase).
    private static readonly HashSet<string> KnownMarkProps = new(StringComparer.OrdinalIgnoreCase)
    {
        "find", "color", "note", "tofix", "regex",
    };

    private static Command BuildMarkCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
        var pathArg = new Argument<string>("path") { Description = "DOM path to the element to mark" };
        var propsOpt = new Option<string[]>("--prop")
        {
            Description = "Mark property: find=..., color=..., note=..., tofix=..., regex=true",
            AllowMultipleArgumentsPerToken = true,
        };

        var cmd = new Command("mark",
            "Attach an in-memory advisory mark to a document element via the watch process. Path must be in data-path format (e.g. /body/p[1]); 'selected' marks all selected elements.");
        cmd.Add(fileArg);
        cmd.Add(pathArg);
        cmd.Add(propsOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var path = result.GetValue(pathArg)!;
            var rawProps = result.GetValue(propsOpt) ?? Array.Empty<string>();

            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string? deprecatedExpectValue = null;
            foreach (var p in rawProps)
            {
                var eq = p.IndexOf('=');
                if (eq <= 0) continue;
                var key = p[..eq];
                var val = p[(eq + 1)..];

                // (a) Deprecated alias: `expect` was renamed to `tofix` in a052fb6.
                // Route the value to `tofix` with a deprecation warning on stderr
                // so old scripts/prompts continue to work instead of silently
                // losing data. Explicit `--prop tofix=...` takes precedence.
                if (string.Equals(key, "expect", StringComparison.OrdinalIgnoreCase))
                {
                    deprecatedExpectValue = val;
                    continue;
                }

                // (c) Unknown prop — warn and ignore instead of dropping silently.
                // This catches typos like --prop noet=... that previously produced
                // a mark with missing fields and no diagnostic.
                if (!KnownMarkProps.Contains(key))
                {
                    Console.Error.WriteLine(
                        $"Warning: unknown property '{key}' for mark, ignored. " +
                        "Known: find, color, note, tofix, regex.");
                    continue;
                }

                props[key] = val;
            }

            if (deprecatedExpectValue != null)
            {
                if (props.ContainsKey("tofix"))
                {
                    // Explicit `tofix` wins — the `expect` value is dropped.
                    // Warn the user the alias was shadowed so they don't wonder
                    // where their value went.
                    Console.Error.WriteLine(
                        "Warning: 'expect' has been renamed to 'tofix'. " +
                        "An explicit 'tofix' was also provided and takes precedence; " +
                        "the 'expect' value was ignored. Please update your scripts.");
                }
                else
                {
                    props["tofix"] = deprecatedExpectValue;
                    Console.Error.WriteLine(
                        "Warning: 'expect' has been renamed to 'tofix'. " +
                        "The value has been applied to 'tofix'. Please update your scripts.");
                }
            }

            // CONSISTENCY(find-regex): 复用 WordHandler.Set.cs:60-61 的 regex→raw-string 转换,
            // 保持 mark 和 set 在 find/regex 词汇上完全一致(literal | r"..." | regex=true flag)。
            // 要修改 find 解析协议,grep "CONSISTENCY(find-regex)" 找全所有调用点项目级一起改,
            // 不要在 mark 单点改。见 CLAUDE.md Design Principles。
            props.TryGetValue("find", out var findText);
            findText ??= "";
            if (props.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag)
                && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
            {
                findText = $"r\"{findText}\"";
            }

            // Build the common prop set once — reused for every target path
            // when the user passes the `selected` pseudo-path.
            var findVal = string.IsNullOrEmpty(findText) ? null : findText;
            var colorVal = props.TryGetValue("color", out var c) ? c : null;
            var noteVal = props.TryGetValue("note", out var n) ? n : null;
            var tofixVal = props.TryGetValue("tofix", out var e) ? e : null;

            // Resolve the target path(s). For the 'selected' pseudo-path, pull the
            // current selection from the running watch process and mark each path
            // individually with the same prop set. Rationale: a block of selected
            // elements is conceptually N independent marks (one per element); a
            // single mark with N paths would need new wire-format plumbing and
            // make find/stale semantics ambiguous.
            List<string> targetPaths;
            if (string.Equals(path, "selected", StringComparison.Ordinal))
            {
                var selection = WatchNotifier.QuerySelection(file.FullName);
                if (selection == null)
                {
                    var err = $"No watch process is running for {file.Name}. Start one with: officecli watch {file.Name}";
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                    else Console.Error.WriteLine(err);
                    return 1;
                }
                if (selection.Length == 0)
                {
                    var err = "No elements are currently selected. Click or drag-select in the watch browser first.";
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                    else Console.Error.WriteLine(err);
                    return 1;
                }
                targetPaths = new List<string>(selection);
            }
            else
            {
                targetPaths = new List<string> { path };
            }

            var createdIds = new List<string>();
            var createdMarks = new List<WatchMark>();
            foreach (var targetPath in targetPaths)
            {
                var req = new MarkRequest
                {
                    Path = targetPath,
                    Find = findVal,
                    Color = colorVal,
                    Note = noteVal,
                    Tofix = tofixVal,
                };

                string? id;
                try
                {
                    id = WatchNotifier.AddMark(file.FullName, req);
                }
                catch (MarkRejectedException rex)
                {
                    // BUG-BT-001: server rejected the request (invalid color, invalid
                    // path, etc.). Surface the actual reason instead of silently
                    // returning success with an empty id.
                    var msg = targetPaths.Count > 1 ? $"{targetPath}: {rex.Message}" : rex.Message;
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
                    else Console.Error.WriteLine(msg);
                    return 1;
                }
                if (id == null)
                {
                    var err = $"No watch process is running for {file.Name}. Start one with: officecli watch {file.Name}";
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                    else Console.Error.WriteLine(err);
                    return 1;
                }
                createdIds.Add(id);
            }

            if (json)
            {
                // Fetch the resolved marks (server has populated matched_text +
                // stale by now) and return them so AI consumers don't need a
                // follow-up get-marks round-trip.
                var full = WatchNotifier.QueryMarksFull(file.FullName);
                if (full != null)
                {
                    var idSet = new HashSet<string>(createdIds);
                    foreach (var m in full.Marks)
                        if (idSet.Contains(m.Id)) createdMarks.Add(m);
                }
                if (createdMarks.Count == targetPaths.Count)
                {
                    if (targetPaths.Count == 1)
                    {
                        var payload = System.Text.Json.JsonSerializer.Serialize(
                            createdMarks[0], WatchMarkJsonOptions.WatchMarkInfo);
                        Console.WriteLine(payload);
                    }
                    else
                    {
                        // Array envelope mirrors MarksResponse shape (no version).
                        var payload = System.Text.Json.JsonSerializer.Serialize(
                            createdMarks.ToArray(), WatchMarkJsonOptions.WatchMarkArrayInfo);
                        Console.WriteLine(payload);
                    }
                }
                else
                {
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(
                        $"Marked {targetPaths.Count} element(s) (ids={string.Join(",", createdIds)})"));
                }
            }
            else
            {
                if (targetPaths.Count == 1)
                    Console.WriteLine($"Marked {targetPaths[0]} (id={createdIds[0]})");
                else
                    Console.WriteLine($"Marked {targetPaths.Count} element(s) (ids={string.Join(",", createdIds)})");
            }
            return 0;
        }, json); });

        return cmd;
    }

    // ==================== unmark ====================

    private static Command BuildUnmarkMarkCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var pathOpt = new Option<string?>("--path") { Description = "Element path to unmark" };
        var allOpt = new Option<bool>("--all") { Description = "Remove all marks for this file" };

        var cmd = new Command("unmark",
            "Remove marks from the watch process. Specify --path <data-path> or --all.");
        cmd.Add(fileArg);
        cmd.Add(pathOpt);
        cmd.Add(allOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var pathVal = result.GetValue(pathOpt);
            var allVal = result.GetValue(allOpt);

            // Require explicit choice — never silently default
            if (allVal && !string.IsNullOrEmpty(pathVal))
            {
                var err = "Specify either --path or --all, not both.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 2;
            }
            if (!allVal && string.IsNullOrEmpty(pathVal))
            {
                var err = "Must specify either --path <p> or --all.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 2;
            }

            var req = new UnmarkRequest { Path = pathVal, All = allVal };
            var removed = WatchNotifier.RemoveMarks(file.FullName, req);
            if (removed == null)
            {
                var err = $"No watch process is running for {file.Name}.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 1;
            }

            var msg = $"Removed {removed} mark(s)";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else Console.WriteLine(msg);
            return 0;
        }, json); });

        return cmd;
    }

    // ==================== get-marks ====================

    private static Command BuildGetMarksCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path" };

        var cmd = new Command("get-marks",
            "List all marks currently held by the watch process.");
        cmd.Add(fileArg);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var full = WatchNotifier.QueryMarksFull(file.FullName);
            if (full == null)
            {
                var err = $"No watch process is running for {file.Name}.";
                // BUG-BT-R4-01: even on error the --json output must keep the
                // {version, marks, error} shape so the SKILL.md jq pipeline
                // (`.marks[] | ...`) doesn't crash with "Cannot iterate over
                // null" when an agent runs the apply pipeline against a dead
                // watch. Empty marks array is the natural "nothing to do" form;
                // the error field carries the human-readable reason. Exit 1
                // still signals failure to script-level checks.
                if (json)
                {
                    // JSON-escape the error message manually to avoid the
                    // reflection-based Serialize<string> overload (IL2026 trim
                    // warning under AOT). The set of chars that actually need
                    // escaping in this context is small.
                    var escaped = err.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
                    var emptyEnvelope = $"{{\"version\":0,\"marks\":[],\"error\":\"{escaped}\"}}";
                    Console.WriteLine(emptyEnvelope);
                }
                else Console.Error.WriteLine(err);
                return 1;
            }

            var marks = full.Marks;

            if (json)
            {
                // Top-level object {version, marks} — no envelope wrapping, no
                // double-encoded JSON-inside-JSON. AI consumers parse once.
                var payload = System.Text.Json.JsonSerializer.Serialize(
                    full, WatchMarkJsonOptions.MarksResponseInfo);
                Console.WriteLine(payload);
            }
            else
            {
                if (marks.Length == 0)
                {
                    Console.WriteLine("(no marks)");
                }
                else
                {
                    Console.WriteLine($"id  path                                              find                  matched  color    note");
                    Console.WriteLine($"--  ------------------------------------------------  --------------------  -------  -------  ----");
                    foreach (var m in marks)
                    {
                        var matchedStr = m.MatchedText.Length == 0
                            ? (m.Stale ? "(stale)" : "-")
                            : (m.MatchedText.Length == 1
                                ? Truncate(m.MatchedText[0], 6)
                                : $"[{string.Join(",", m.MatchedText.Take(2).Select(t => Truncate(t, 4)))}]({m.MatchedText.Length})");
                        Console.WriteLine($"{m.Id,-3} {Truncate(m.Path, 48),-48}  {Truncate(m.Find ?? "-", 20),-20}  {matchedStr,-7}  {Truncate(m.Color ?? "-", 7),-7}  {Truncate(m.Note ?? "-", 30)}");
                    }
                }
            }
            return 0;
        }, json); });

        return cmd;
    }

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s.Substring(0, max - 1) + "…";
}
