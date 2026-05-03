// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildSetCommand(Option<bool> jsonOption)
    {
        var forceOption = new Option<bool>("--force") { Description = "Force write even if document is protected" };
        var setFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var setPathArg = new Argument<string>("path") { Description = "DOM path to the element" };
        var propsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value)", AllowMultipleArgumentsPerToken = true };

        var setCommand = new Command("set", "Modify a document node's properties") { TreatUnmatchedTokensAsErrors = false };
        setCommand.Add(setFileArg);
        setCommand.Add(setPathArg);
        setCommand.Add(propsOpt);
        setCommand.Add(jsonOption);
        setCommand.Add(forceOption);

        setCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(setFileArg)!;
            var path = result.GetValue(setPathArg)!;
            var props = result.GetValue(propsOpt);
            var force = result.GetValue(forceOption);

            // BUG-BT-R5-01: support the `selected` pseudo-path (mark and get
            // already do). Expand to the first selected path and recursively
            // re-invoke set for any additional paths after the main set
            // completes. CONSISTENCY(selected-pseudo): grep for the same
            // pseudo-path handling in CommandBuilder.Mark.cs / GetQuery.cs.
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

            // Path that does not start with '/' is rejected up front (must run before
            // TryResident — resident has its own dispatch and would otherwise execute
            // selector-mode silently). handler.Set treats no-slash paths as CSS selectors
            // (Query→Set per match), so a typo like "section[1]" would corrupt the doc with
            // no way for the user to notice. Selector-mode is opt-in via the `query`
            // subcommand, not via dropping the slash. CONSISTENCY(no-slash-reject):
            // ResidentServer.ExecuteSet enforces the same rule.
            if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
            {
                var err = $"Error: path '{path}' must start with '/'. Did you mean '/{path}'?";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 1;
            }

            // CONSISTENCY(prop-key-case): --prop keys are case-insensitive
            // so "SRC=x" and "src=x" both resolve to the same handler key.
            // Reuse ParsePropsArray so the inline and resident-server paths
            // stay in sync.
            var properties = ParsePropsArray(props);

            var extension = Path.GetExtension(file.FullName);
            if (string.Equals(extension, ".hwp", StringComparison.OrdinalIgnoreCase)
                && string.Equals(path, "/field", StringComparison.OrdinalIgnoreCase))
                return HandleHwpFieldSet(file.FullName, HwpFormat.Hwp, properties, json);
            if (string.Equals(extension, ".hwp", StringComparison.OrdinalIgnoreCase)
                && string.Equals(path, "/text", StringComparison.OrdinalIgnoreCase))
                return HandleHwpTextReplace(file.FullName, HwpFormat.Hwp, properties, json);
            if (string.Equals(extension, ".hwp", StringComparison.OrdinalIgnoreCase)
                && string.Equals(path, "/table/cell", StringComparison.OrdinalIgnoreCase))
                return HandleHwpTableCellSet(file.FullName, HwpFormat.Hwp, properties, json);

            if (string.Equals(extension, ".hwpx", StringComparison.OrdinalIgnoreCase)
                && HwpEngineSelector.IsExperimentalBridgeEnabled()
                && string.Equals(path, "/field", StringComparison.OrdinalIgnoreCase))
                return HandleHwpFieldSet(file.FullName, HwpFormat.Hwpx, properties, json);
            if (string.Equals(extension, ".hwpx", StringComparison.OrdinalIgnoreCase)
                && HwpEngineSelector.IsExperimentalBridgeEnabled()
                && string.Equals(path, "/text", StringComparison.OrdinalIgnoreCase))
                return HandleHwpTextReplace(file.FullName, HwpFormat.Hwpx, properties, json);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "set";
                req.Args["path"] = path;
                req.Props = properties;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);

            var unsupported = handler.Set(path, properties);

            // Scope the unsupported-prop fuzzy-suggestion pool by handler type
            // so e.g. Excel pivot errors don't suggest PPTX-only keys like
            // 'rotation' for an unknown 'location' prop (R2-4).
            string? suggestionScope = handler switch
            {
                OfficeCli.Handlers.ExcelHandler => "excel",
                OfficeCli.Handlers.WordHandler => "word",
                OfficeCli.Handlers.PowerPointHandler => "pptx",
                _ => null,
            };

            // Auto-correct: attempt to fix unsupported properties with Levenshtein distance == 1
            var autoCorrected = new List<(string Original, string Corrected, string Value)>();
            var stillUnsupported = new List<string>();
            foreach (var u in unsupported)
            {
                var rawKey = u.Contains(' ') ? u[..u.IndexOf(' ')] : u;
                if (properties.TryGetValue(rawKey, out var val))
                {
                    var (suggestion, dist, isUnique) = SuggestPropertyWithDistance(rawKey, suggestionScope);
                    if (suggestion != null && dist == 1 && isUnique)
                    {
                        // Auto-correct: re-apply with corrected key
                        var correctedProps = new Dictionary<string, string> { [suggestion] = val };
                        var retryUnsupported = handler.Set(path, correctedProps);
                        if (retryUnsupported.Count == 0)
                        {
                            autoCorrected.Add((rawKey, suggestion, val));
                            continue;
                        }
                    }
                }
                stillUnsupported.Add(u);
            }

            // unsupported entries may contain help text like "key (valid props: ...)" — extract raw keys
            var unsupportedKeys = stillUnsupported.Select(u => u.Contains(' ') ? u[..u.IndexOf(' ')] : u).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var autoCorrectedKeys = autoCorrected.Select(ac => ac.Original).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var applied = properties.Where(kv => !unsupportedKeys.Contains(kv.Key) && !autoCorrectedKeys.Contains(kv.Key)).ToList();
            // Include auto-corrected props in applied list with the corrected key name
            foreach (var ac in autoCorrected)
                applied.Add(new KeyValuePair<string, string>(ac.Corrected, ac.Value));

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

            var message = applied.Count > 0
                ? $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
                  + (findMatchCount.HasValue ? $" ({findMatchCount.Value} matched)" : "")
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

            if (json)
            {
                var allWarnings = new List<OfficeCli.Core.CliWarning>();
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
                    var suggestion = SuggestPropertyScoped(p, suggestionScope);
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
                var outputMsg = setSpatialLine != null ? $"{message}\n  {setSpatialLine}" : message;
                bool allFailed = applied.Count == 0 && (stillUnsupported.Count > 0 || unsupported.Count > 0);
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
            return 0;
        }, json); });

        return setCommand;
    }

    private static int HandleHwpFieldSet(
        string inputPath,
        HwpFormat format,
        Dictionary<string, string> properties,
        bool json)
    {
        var fieldName = FirstValue(properties, "name", "field", "field-name");
        var fieldIdRaw = FirstValue(properties, "id", "field-id", "fieldId");
        var value = FirstValue(properties, "value", "text");
        var output = FirstValue(properties, "output", "out");
        if ((string.IsNullOrWhiteSpace(fieldName) && string.IsNullOrWhiteSpace(fieldIdRaw))
            || value == null || string.IsNullOrWhiteSpace(output))
        {
            var message = "HWP field set requires --prop name=<field> or --prop id=<fieldId>, plus --prop value=<text> --prop output=<path>.";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(message));
            else Console.Error.WriteLine(message);
            return 1;
        }
        if (!string.IsNullOrWhiteSpace(fieldIdRaw) && !int.TryParse(fieldIdRaw, out _))
        {
            var message = $"Invalid HWP field id '{fieldIdRaw}'.";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(message));
            else Console.Error.WriteLine(message);
            return 1;
        }

        var outputPath = Path.GetFullPath(output);
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir)) Directory.CreateDirectory(outputDir);

        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
        var engine = HwpEngineSelector.GetEngine(formatKey, HwpCapabilityConstants.OperationFillField);
        var nameFields = string.IsNullOrWhiteSpace(fieldName)
            ? new Dictionary<string, string>()
            : new Dictionary<string, string> { [fieldName] = value };
        var idFields = string.IsNullOrWhiteSpace(fieldIdRaw)
            ? new Dictionary<int, string>()
            : new Dictionary<int, string> { [int.Parse(fieldIdRaw)] = value };
        var request = new HwpFillFieldRequest(
            format,
            inputPath,
            outputPath,
            nameFields,
            json)
        {
            FieldIds = idFields
        };
        var result = engine.FillFieldAsync(request, CancellationToken.None).GetAwaiter().GetResult();
        var fieldLabel = !string.IsNullOrWhiteSpace(fieldName) ? fieldName : $"#{fieldIdRaw}";

        if (json)
        {
            var envelope = new System.Text.Json.Nodes.JsonObject
            {
                ["success"] = true,
                ["message"] = $"Updated HWP field '{fieldLabel}' -> {result.OutputPath}",
                ["data"] = new System.Text.Json.Nodes.JsonObject
                {
                    ["outputPath"] = result.OutputPath,
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["evidence"] = HwpCapabilityJsonMapper.ToJsonArray(result.Evidence)
                },
                ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
            };
            Console.WriteLine(envelope.ToJsonString(OutputFormatter.PublicJsonOptions));
        }
        else
        {
            Console.WriteLine($"Updated HWP field '{fieldLabel}' -> {result.OutputPath}");
            foreach (var warning in result.Warnings)
                Console.Error.WriteLine($"WARNING: {warning}");
        }
        return 0;
    }

    private static int HandleHwpTextReplace(
        string inputPath,
        HwpFormat format,
        Dictionary<string, string> properties,
        bool json)
    {
        var query = FirstValue(properties, "find", "query", "old");
        var value = FirstValue(properties, "value", "text", "new");
        var output = FirstValue(properties, "output", "out");
        var mode = FirstValue(properties, "mode") ?? "one";
        var caseSensitiveRaw = FirstValue(properties, "case-sensitive", "caseSensitive") ?? "false";
        if (string.IsNullOrEmpty(query) || value == null || string.IsNullOrWhiteSpace(output))
        {
            var message = "HWP text replace requires --prop find=<text> --prop value=<text> --prop output=<path>.";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(message));
            else Console.Error.WriteLine(message);
            return 1;
        }

        var outputPath = Path.GetFullPath(output);
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir)) Directory.CreateDirectory(outputDir);

        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
        var engine = HwpEngineSelector.GetEngine(formatKey, HwpCapabilityConstants.OperationReplaceText);
        var request = new HwpReplaceTextRequest(
            format,
            inputPath,
            outputPath,
            query,
            value,
            mode,
            bool.TryParse(caseSensitiveRaw, out var caseSensitive) && caseSensitive,
            json);
        var result = engine.ReplaceTextAsync(request, CancellationToken.None).GetAwaiter().GetResult();

        if (json)
        {
            var envelope = new System.Text.Json.Nodes.JsonObject
            {
                ["success"] = true,
                ["message"] = $"Replaced HWP text '{query}' -> {result.OutputPath}",
                ["data"] = new System.Text.Json.Nodes.JsonObject
                {
                    ["outputPath"] = result.OutputPath,
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["evidence"] = HwpCapabilityJsonMapper.ToJsonArray(result.Evidence)
                },
                ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
            };
            Console.WriteLine(envelope.ToJsonString(OutputFormatter.PublicJsonOptions));
        }
        else
        {
            Console.WriteLine($"Replaced HWP text '{query}' -> {result.OutputPath}");
            foreach (var warning in result.Warnings)
                Console.Error.WriteLine($"WARNING: {warning}");
        }
        return 0;
    }

    private static string? FirstValue(Dictionary<string, string> properties, params string[] keys)
    {
        foreach (var key in keys)
            if (properties.TryGetValue(key, out var value))
                return value;
        return null;
    }
}
