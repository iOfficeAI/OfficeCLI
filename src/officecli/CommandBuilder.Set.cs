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
        var propsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value)", AllowMultipleArgumentsPerToken = true };
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

            // Scope the unsupported-prop fuzzy-suggestion pool by handler type
            // so e.g. Excel pivot errors don't suggest PPTX-only keys like
            // 'rotation' for an unknown 'location' prop (R2-4). Kept local: the
            // shared core computes its own copy, but the CLI warning / extra-path
            // decoration below also needs it.
            string? suggestionScope = handler switch
            {
                OfficeCli.Handlers.ExcelHandler => "excel",
                OfficeCli.Handlers.WordHandler => "word",
                OfficeCli.Handlers.PowerPointHandler => "pptx",
                _ => null,
            };

            // Shared core: apply + prop-autocorrect + categorise (one copy for
            // CLI / batch / MCP / resident; see ApplySetWithCorrection). The rich
            // CLI envelope below — find-count, position overlap, --json warnings,
            // exit codes — stays here.
            // CONSISTENCY(applied-echo): pre/post Format snapshots feed the
            // " (applied: ...)" normalization echo; mirrored in
            // ResidentServer.ExecuteSet.
            var beforeSnap = TryGetFormatSnapshot(handler, path);
            var (applied, stillUnsupported, autoCorrected) = ApplySetWithCorrection(handler, path, properties);

            // Get find match count if applicable.
            // CONSISTENCY(find-match-count): mirrored in ResidentServer.ExecuteSet.
            // The resident path is hit whenever a resident process is open
            // (which `create` does by default), so both sites must surface
            // findMatchCount + zero_matches warning identically.
            int? findMatchCount = null;
            if (properties.ContainsKey("find"))
            {
                findMatchCount = handler switch
                {
                    OfficeCli.Handlers.WordHandler wh => wh.LastFindMatchCount,
                    OfficeCli.Handlers.PowerPointHandler ph => ph.LastFindMatchCount,
                    OfficeCli.Handlers.ExcelHandler eh => eh.LastFindMatchCount,
                    _ => null
                };
            }

            // CONSISTENCY(selector-set): echo how many elements a multi-match
            // selector set touched so a Sheet1!row[工资>5000]-style change shows
            // its scope. Single-target paths (count 1) stay quiet.
            int? selectorCount = !isSelectorSet ? null : handler switch
            {
                OfficeCli.Handlers.WordHandler wh => wh.LastSelectorSetCount,
                OfficeCli.Handlers.PowerPointHandler ph => ph.LastSelectorSetCount,
                OfficeCli.Handlers.ExcelHandler eh => eh.LastSelectorSetCount,
                _ => null
            };

            // R4-bt-1: an equation mode switch MOVES the element (oMathPara ⇄
            // oMath), changing its canonical path. Report the NEW resolvable
            // path so the "Updated …" line points at a path that still resolves.
            var reportPath = (handler as OfficeCli.Handlers.WordHandler)?.LastSetNewPath ?? path;
            var appliedSuffix = BuildAppliedSuffix(applied,
                beforeSnap, TryGetFormatSnapshot(handler, reportPath));
            var message = applied.Count > 0
                ? $"Updated {reportPath}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
                  + appliedSuffix
                  + (findMatchCount.HasValue ? $" ({findMatchCount.Value} matched)" : "")
                  + (selectorCount > 1 ? $" ({selectorCount} elements matched)" : "")
                : $"Error: No properties applied to {path}";

            // Check if position-related props were changed → show coordinates + overlap warning
            var positionChanged = applied.Any(kv => PositionKeys.Contains(kv.Key));
            string? setSpatialLine = null;
            var setOverlaps = new List<string>();
            if (positionChanged)
            {
                setSpatialLine = GetPptSpatialLine(handler, path);
                if (setSpatialLine != null) setOverlaps = CheckPositionOverlap(handler, path);
            }

            // Unrecognized LaTeX commands/environments from an equation Set
            // (formula=). Same UX as unsupported_property (warning + JSON
            // envelope + exit 2); the equation is still written (lenient
            // accept). CONSISTENCY: mirrors CommandBuilder.Add and
            // ResidentServer.ExecuteSet.
            var setUnrecognizedLatex = handler switch
            {
                OfficeCli.Handlers.WordHandler wlx => wlx.LastUnrecognizedLatex,
                OfficeCli.Handlers.PowerPointHandler plx => plx.LastUnrecognizedLatex,
                _ => null,
            };
            bool hasUnrecognizedLatex = setUnrecognizedLatex is { Count: > 0 };

            if (json)
            {
                var allWarnings = new List<OfficeCli.Core.CliWarning>();
                if (hasUnrecognizedLatex)
                {
                    foreach (var tok in setUnrecognizedLatex!)
                        allWarnings.Add(new OfficeCli.Core.CliWarning
                        {
                            Message = $"unrecognized_latex_command: {tok}",
                            Code = "unrecognized_latex_command",
                            Suggestion = "Check the command spelling; see https://katex.org/docs/supported.html for supported syntax.",
                        });
                }
                if (findMatchCount is 0)
                {
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"find pattern matched 0 occurrences at {path} — original text may have been edited or the path is wrong",
                        Code = "zero_matches",
                        Suggestion = "verify the path still resolves and the find text is current"
                    });
                }
                foreach (var ac in autoCorrected)
                {
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"Auto-corrected '{ac.Original}' to '{ac.Corrected}'",
                        Code = "auto_corrected",
                        Suggestion = ac.Corrected
                    });
                }
                foreach (var p in stillUnsupported)
                {
                    // An entry that already carries a handler-embedded hint
                    // ("薪水 (no such column; available: …)") must not get a
                    // generic did-you-mean stacked on top — the handler already
                    // said what's valid. Mirrors CommandBuilder.FormatUnsupported.
                    var suggestion = p.Contains('(') ? null : SuggestPropertyScoped(p, suggestionScope);
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = suggestion != null ? $"Unsupported property: {p} (did you mean: {suggestion}?)" : $"Unsupported property: {p}",
                        Code = "unsupported_property",
                        Suggestion = suggestion
                    });
                }
                if (setOverlaps.Count > 0)
                {
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"Same position as {string.Join(", ", setOverlaps)}",
                        Code = "position_overlap",
                        Suggestion = "Use different x/y values to avoid overlap"
                    });
                }
                var setOverflow = CheckTextOverflow(handler, path);
                if (setOverflow != null)
                {
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = setOverflow,
                        Code = "text_overflow",
                        Suggestion = "Increase shape height/width, reduce font size, or shorten text"
                    });
                }
                if (handler is OfficeCli.Handlers.WordHandler setWhWarn)
                {
                    foreach (var w in setWhWarn.LastSetWarnings)
                        allWarnings.Add(new OfficeCli.Core.CliWarning { Message = w, Code = "advisory" });
                }
                var outputMsg = setSpatialLine != null ? $"{message}\n  {setSpatialLine}" : message;
                // applied==0 implies no key auto-corrected (corrections land in
                // applied), so stillUnsupported already equals the raw set, and
                // the old `|| unsupported.Count>0` term was redundant.
                bool allFailed = applied.Count == 0 && stillUnsupported.Count > 0;
                Console.WriteLine(allFailed
                    ? OutputFormatter.WrapEnvelopeError(outputMsg, allWarnings.Count > 0 ? allWarnings : null)
                    : OutputFormatter.WrapEnvelopeText(outputMsg, allWarnings.Count > 0 ? allWarnings : null, findMatchCount));
            }
            else
            {
                foreach (var ac in autoCorrected)
                    Console.Error.WriteLine($"WARNING: Auto-corrected '{ac.Original}' to '{ac.Corrected}'");
                Console.WriteLine(message);
                if (findMatchCount is 0)
                    Console.Error.WriteLine($"WARNING: find pattern matched 0 occurrences at {path}");
                if (setSpatialLine != null) Console.WriteLine($"  {setSpatialLine}");
                if (setOverlaps.Count > 0)
                    Console.Error.WriteLine($"  WARNING: Same position as {string.Join(", ", setOverlaps)}");
                var setOverflowPlain = CheckTextOverflow(handler, path);
                if (setOverflowPlain != null)
                    Console.Error.WriteLine($"  WARNING: {setOverflowPlain}");
                if (stillUnsupported.Count > 0)
                    Console.Error.WriteLine(FormatUnsupported(stillUnsupported, suggestionScope));
                if (handler is OfficeCli.Handlers.WordHandler setWhWarnPlain)
                {
                    foreach (var w in setWhWarnPlain.LastSetWarnings)
                        Console.Error.WriteLine($"  WARNING: {w}");
                }
                if (hasUnrecognizedLatex)
                    foreach (var tok in setUnrecognizedLatex!)
                        Console.Error.WriteLine($"  WARNING: unrecognized_latex_command: {tok}");
            }
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

            if (stillUnsupported.Count > 0) return 2;
            if (hasUnrecognizedLatex) return 2;
            return 0;
        }, json); });

        return setCommand;
    }
}
