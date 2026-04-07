// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    // ==================== mark ====================

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
            "Attach an in-memory advisory mark to a document element via the running watch process. " +
            "Marks are not written to the file. " +
            "Path must be in data-path format (e.g. /p[1], /slide[1]/shape[@id=N]), as emitted by watch HTML preview. " +
            "Inspect the rendered HTML for valid paths. Native handler query paths like /body/p[@paraId=...] will not resolve.");
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
            foreach (var p in rawProps)
            {
                var eq = p.IndexOf('=');
                if (eq > 0) props[p[..eq]] = p[(eq + 1)..];
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

            var req = new MarkRequest
            {
                Path = path,
                Find = string.IsNullOrEmpty(findText) ? null : findText,
                Color = props.TryGetValue("color", out var c) ? c : null,
                Note = props.TryGetValue("note", out var n) ? n : null,
                Tofix = props.TryGetValue("tofix", out var e) ? e : null,
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
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(rex.Message));
                else Console.Error.WriteLine(rex.Message);
                return 1;
            }
            if (id == null)
            {
                var err = $"No watch process is running for {file.Name}. Start one with: officecli watch {file.Name}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 1;
            }

            if (json)
            {
                // Fetch the resolved mark (server has populated matched_text +
                // stale by now) and return the full WatchMark object so AI
                // consumers don't need a follow-up get-marks round-trip.
                var full = WatchNotifier.QueryMarksFull(file.FullName);
                WatchMark? resolved = null;
                if (full != null)
                {
                    for (int i = 0; i < full.Marks.Length; i++)
                    {
                        if (full.Marks[i].Id == id) { resolved = full.Marks[i]; break; }
                    }
                }
                if (resolved != null)
                {
                    var payload = System.Text.Json.JsonSerializer.Serialize(
                        resolved, WatchMarkJsonOptions.WatchMarkInfo);
                    Console.WriteLine(payload);
                }
                else
                {
                    // Fallback: only the id is guaranteed. Shouldn't happen in
                    // practice because the add-then-query sequence races only
                    // with unmark, which CLI doesn't do here.
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText($"Marked {path} (id={id})"));
                }
            }
            else
            {
                Console.WriteLine($"Marked {path} (id={id})");
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
            "Remove marks from the running watch process. Must specify either --path or --all. " +
            "--path must be in data-path format (e.g. /p[1], /slide[1]/shape[@id=N]), matching the value used with mark. " +
            "Native handler query paths like /body/p[@paraId=...] will not match.");
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
            "List all marks currently held by the running watch process. " +
            "Paths in the output are in data-path format (e.g. /p[1], /slide[1]/shape[@id=N]), " +
            "not native handler query paths.");
        cmd.Add(fileArg);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var full = WatchNotifier.QueryMarksFull(file.FullName);
            if (full == null)
            {
                var err = $"No watch process is running for {file.Name}.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
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
