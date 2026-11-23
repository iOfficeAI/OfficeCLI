// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildSetCommand(Option<bool> jsonOption)
    {
        var forceOption = new Option<bool>("--force") { Description = "Force write even if document is protected" };
        var setFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var setPathArg = new Argument<string>("path") { Description = "DOM path to the element. The 'selected' pseudo-path is deprecated for mutations: use `get selected` to capture path(s) first, then `set <path>` (or a `batch` file for multi-select) so the target lives in the command line, not in transient watch-server state." };
        var propsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value). xlsx range paths accept stylepreset=table_header|table_banded|kpi_card|metric_positive|metric_negative|note_gray (free-range styling; ListObject tables use add --type table --prop style=…)", AllowMultipleArgumentsPerToken = true };
        // Selector: top-level alternative to --prop find=VALUE. r"..." prefix triggers regex (project-wide CONSISTENCY(find-regex)).
        var findOpt = new Option<string?>("--find") { Description = "Find this text/pattern (literal substring; `r\"...\"` prefix enables regex). Equivalent to --prop find=VALUE." };
        // Action paired with --find: replacement text. Top-level alternative to --prop replace=VALUE.
        var replaceOpt = new Option<string?>("--replace") { Description = "Replacement text for --find matches. Equivalent to --prop replace=VALUE." };

        var setCommand = new Command("set", "Modify a document node's properties") { TreatUnmatchedTokensAsErrors = false };
        setCommand.Add(setFileArg);
        setCommand.Add(setPathArg);
        setCommand.Add(propsOpt);
        setCommand.Add(findOpt);
        setCommand.Add(replaceOpt);
        setCommand.Add(jsonOption);
        setCommand.Add(forceOption);

        setCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            // JSON mode: collect Core-layer advisory warnings (sites too deep
            // to reach this command's local warning lists, e.g. the number-
            // format check in ExcelStyleManager) so WrapEnvelope* folds them
            // into warnings[]. CONSISTENCY(numfmt-warning): resident-routed
            // commands get the same via ResidentServer.BuildWarnings.
            if (json) OfficeCli.Core.WarningContext.Begin();
            var file = result.GetValue(setFileArg)!;
            var path = MsysPathHint.Restore(result.GetValue(setPathArg)!)!;
            var props = result.GetValue(propsOpt);
            // TreatUnmatchedTokensAsErrors=false lets a bare key=value be
            // reported as a missing --prop; anything else that is unmatched is
            // an unknown option and must not vanish with exit 0 (add already
            // does both — set did neither).
            var setUnmatchedKv = DetectUnmatchedKeyValues(result);
            foreach (var kv in setUnmatchedKv)
                Console.Error.WriteLine($"WARNING: Bare property '{kv}' ignored. Did you mean: --prop {kv}");
            RejectUnknownOptionTokens(result, setUnmatchedKv);
            var findFlag = result.GetValue(findOpt);
            var replaceFlag = result.GetValue(replaceOpt);
            var force = result.GetValue(forceOption);

            // Selector-flag migration: --find / --replace are the canonical
            // top-level forms; --prop find=VALUE / --prop replace=VALUE remain
            // accepted but emit a deprecation hint. Merging the flags into the
            // props array (as "find=<value>" / "replace=<value>") keeps every
            // downstream consumer (handlers, resident server, batch) working
            // with zero changes — the flags are pure syntactic sugar.
            var hasPropFind = props?.Any(p => p.StartsWith("find=", StringComparison.OrdinalIgnoreCase)) == true;
            var hasPropReplace = props?.Any(p => p.StartsWith("replace=", StringComparison.OrdinalIgnoreCase)) == true;
            if (findFlag != null && hasPropFind)
            {
                var err = "Cannot combine --find and --prop find=. Use --find only.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine($"Error: {err}");
                return 1;
            }
            if (replaceFlag != null && hasPropReplace)
            {
                var err = "Cannot combine --replace and --prop replace=. Use --replace only.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine($"Error: {err}");
                return 1;
            }
            if (findFlag != null || replaceFlag != null)
            {
                var merged = props?.ToList() ?? new List<string>();
                if (findFlag != null) merged.Add($"find={findFlag}");
                if (replaceFlag != null) merged.Add($"replace={replaceFlag}");
                props = merged.ToArray();
            }
            else if (hasPropFind || hasPropReplace)
            {
                var legacy = hasPropFind && hasPropReplace ? "find / replace"
                           : hasPropFind ? "find" : "replace";
                var flag   = hasPropFind && hasPropReplace ? "--find / --replace"
                           : hasPropFind ? "--find" : "--replace";
                Console.Error.WriteLine($"Hint: prefer `{flag} VALUE` over `--prop {legacy}=VALUE` (selector/action keys are migrating out of --prop).");
            }

            // BUG-BT-R5-01: support the `selected` pseudo-path (mark and get
            // already do). Expand to the first selected path and recursively
            // re-invoke set for any additional paths after the main set
            // completes. CONSISTENCY(selected-pseudo): grep for the same
            // pseudo-path handling in CommandBuilder.Mark.cs / GetQuery.cs.
            //
            // Discouraged for mutations, single- or multi-select alike.
            // `set selected` resolves the selection at *execution* time, so
            // if the user clicks a different element between deciding-to-set
            // and the command running, this branch silently retargets the
            // new element — no error, just the wrong object mutated. The
            // canonical pattern is two-step: `get selected` to freeze the
            // path(s) as strings, then issue explicit `set <path>` per path
            // (or a `batch` file for multi-select). Every mutation command
            // then carries its target in the command line itself — auditable,
            // diffable, replayable from shell history. Out-of-band watch-server
            // state must not influence what mutation commands target. Do not
            // extend this pseudo-path to more mutation commands.
            List<string>? extraSelectedPaths = null;
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
                path = selection[0];
                if (selection.Length > 1)
                {
                    extraSelectedPaths = new List<string>(selection.Length - 1);
                    for (int i = 1; i < selection.Length; i++) extraSelectedPaths.Add(selection[i]);
                }
            }

            // Check document protection for .docx files
            // Skip protection check if the user is changing the protection mode itself
            var isProtectionChange = props?.Any(p => p.StartsWith("protection=", StringComparison.OrdinalIgnoreCase)) == true;
            if (!force && !isProtectionChange && file.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                var protectionError = CheckDocxProtection(file.FullName, path, json);
                if (protectionError != 0) return protectionError;
            }

            // Detect bare key=value positional arguments (missing --prop)
            var unmatchedKvWarnings = DetectUnmatchedKeyValues(result);
            if (unmatchedKvWarnings.Count > 0)
            {
                if (json)
                {
                    var kvWarnings = unmatchedKvWarnings.Select(kv => new OfficeCli.Core.CliWarning
                    {
                        Message = $"Bare property '{kv}' ignored. Use --prop {kv}",
                        Code = "missing_prop_flag",
                        Suggestion = $"--prop {kv}"
                    }).ToList();
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(
                        $"Properties specified without --prop flag. Use: officecli set <file> <path> --prop {string.Join(" --prop ", unmatchedKvWarnings)}",
                        kvWarnings));
                }
                else
                {
                    foreach (var kv in unmatchedKvWarnings)
                        Console.Error.WriteLine($"WARNING: Bare property '{kv}' ignored. Did you mean: --prop {kv}");
                    Console.Error.WriteLine("Hint: Properties must be passed with --prop flag, e.g. officecli set <file> <path> --prop key=value");
                }
                if (props == null || props.Length == 0)
                    return 2;
            }

            // `set` is a mutation command; an invocation with no --prop and
            // no bare unmatched key=value pairs is a caller mistake (missing
            // the property they intended to apply). Without this guard the
            // command ran the dispatcher with an empty properties dictionary,
            // applied nothing, and returned "Updated <path>" with success=0
            // — indistinguishable from a real successful set. Surface as
            // missing_property so the caller knows nothing happened.
            if ((props == null || props.Length == 0) && unmatchedKvWarnings.Count == 0)
            {
                var err = $"No properties to set at {path}. Pass at least one --prop key=value.";
                if (json)
                {
                    var missingPropWarnings = new List<OfficeCli.Core.CliWarning>
                    {
                        new() { Message = err, Code = "missing_property", Suggestion = "--prop key=value" }
                    };
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(err, missingPropWarnings));
                }
                else
                {
                    Console.Error.WriteLine($"Error: {err}");
                }
                return 1;
            }

            // Selector / Excel-native paths (not starting with '/') are accepted:
            // handler.Set treats them as a Query→Set-per-match selector, the same
            // engine `batch` uses, so `set` is consistent with get/query which
            // already take both the XPath form (/Sheet1/A1) and the Excel form
            // (Sheet1!A1, Sheet1!row[工资>5000]). Safety rests on the handler
            // selector branch throwing on an empty match (no silent no-op) and the
            // match-count echo below making a multi-element change visible.
            // CONSISTENCY(selector-set): ResidentServer.ExecuteSet mirrors this.
            var isSelectorSet = !string.IsNullOrEmpty(path) && !path.StartsWith("/");

            // Agent-safety: reject a bare unscoped selector (`cell[...]`, `run`) —
            // it would mutate across the whole document. Allows `/`-scoped paths
            // and Excel `Sheet1!A1`. query is unaffected. Runs before TryResident
            // so the resident-forward path is guarded too.
            OfficeCli.Core.MutationSelectorGuard.EnsureScoped(path, "set");

            if (TryResident(file.FullName, req =>
            {
                req.Command = "set";
                req.Args["path"] = path;
                req.Props = ParsePropsArray(props);
            }, json) is {} rc) return rc;

            // CONSISTENCY(prop-key-case): --prop keys are case-insensitive
            // so "SRC=x" and "src=x" both resolve to the same handler key.
            // Reuse ParsePropsArray so the inline and resident-server paths
            // stay in sync.
            var properties = ParsePropsArray(props);

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            using var deferredOut = OfficeCli.Core.DeferredStdout.Begin(handler); // #493: print only after the save lands

            var outcome = RunSet(handler, file.FullName, path, properties, json, isSelectorSet);
            var stillUnsupported = outcome.Unsupported;
            var suggestionScope = outcome.SuggestionScope;
            NotifyWatch(handler, file.FullName, path);

            // BUG-BT-R5-01: apply the same prop set to the remaining selected
            // paths. Each call goes through handler.Set independently so each
            // path gets its own auto-correct, find-count, and unsupported list,
            // matching the per-path semantics that mark already uses for
            // `mark <file> selected`. We collect any non-zero return as an
            // error escalation but keep going so partial application is at
            // least observable.
            if (extraSelectedPaths is not null && extraSelectedPaths.Count > 0)
            {
                var extraStillUnsupported = false;
                foreach (var extraPath in extraSelectedPaths)
                {
                    var extraResult = handler.Set(extraPath, properties);
                    if (extraResult.Count > 0)
                    {
                        extraStillUnsupported = true;
                        if (!json)
                            Console.Error.WriteLine($"  {extraPath}: {FormatUnsupported(extraResult, suggestionScope)}");
                    }
                    NotifyWatch(handler, file.FullName, extraPath);
                }
                if (extraStillUnsupported && stillUnsupported.Count == 0) return 2;
            }

            return outcome.ExitCode;
        }, json); });

        return setCommand;
    }
}
