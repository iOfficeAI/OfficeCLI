// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildAddCommand(Option<bool> jsonOption)
    {
        var addFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var addParentPathArg = new Argument<string>("parent") { Description = "Parent DOM path (e.g. /body, /Sheet1, /slide[1])" };
        var addTypeOpt = new Option<string>("--type") { Description = "Element type to add (e.g. paragraph, run, table, sheet, row, cell, slide, shape, picture, ole, video)" };
        var addFromOpt = new Option<string?>("--from") { Description = "Copy from an existing element path (e.g. /slide[1]/shape[2])" };
        var addIndexOpt = new Option<int?>("--index") { Description = "Insert position (0-based). If omitted, appends to end" };
        var addAfterOpt = new Option<string?>("--after") { Description = "Insert after the element at this path (e.g. p[@paraId=1A2B3C4D])" };
        var addBeforeOpt = new Option<string?>("--before") { Description = "Insert before the element at this path" };
        var addPropsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value, e.g. --prop src=image.png --prop width=6in)", AllowMultipleArgumentsPerToken = true };
        var forceOption = new Option<bool>("--force") { Description = "Force write even if document is protected" };

        var addCommand = new Command("add", "Add a new element to the document") { TreatUnmatchedTokensAsErrors = false };
        addCommand.Add(addFileArg);
        addCommand.Add(addParentPathArg);
        addCommand.Add(addTypeOpt);
        addCommand.Add(addFromOpt);
        addCommand.Add(addIndexOpt);
        addCommand.Add(addAfterOpt);
        addCommand.Add(addBeforeOpt);
        addCommand.Add(addPropsOpt);
        addCommand.Add(jsonOption);
        addCommand.Add(forceOption);

        addCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(addFileArg)!;
            var parentPath = result.GetValue(addParentPathArg)!;
            var type = result.GetValue(addTypeOpt);
            var from = result.GetValue(addFromOpt);
            var index = result.GetValue(addIndexOpt);
            var after = result.GetValue(addAfterOpt);
            var before = result.GetValue(addBeforeOpt);
            var props = result.GetValue(addPropsOpt);
            var force = result.GetValue(forceOption);

            // Validate mutual exclusivity of --index, --after, --before
            var posCount = (index.HasValue ? 1 : 0) + (after != null ? 1 : 0) + (before != null ? 1 : 0);
            if (posCount > 1)
                throw new OfficeCli.Core.CliException("--index, --after, and --before are mutually exclusive. Use only one.")
                {
                    Code = "invalid_argument",
                    Suggestion = "Use --index for positional insert, or --after/--before for anchor-based insert."
                };

            InsertPosition? position = index.HasValue ? InsertPosition.AtIndex(index.Value)
                : after != null ? InsertPosition.AfterElement(after)
                : before != null ? InsertPosition.BeforeElement(before)
                : null;
            bool hadWarnings = false;

            // Check document protection for .docx files
            if (!force && file.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                var protectionError = CheckDocxProtection(file.FullName, parentPath, json);
                if (protectionError != 0) return protectionError;
            }

            // Detect bare key=value positional arguments (missing --prop)
            var unmatchedKvWarnings = DetectUnmatchedKeyValues(result);
            if (unmatchedKvWarnings.Count > 0)
            {
                hadWarnings = true;
                if (json)
                {
                    var kvWarnings = unmatchedKvWarnings.Select(kv => new OfficeCli.Core.CliWarning
                    {
                        Message = $"Bare property '{kv}' ignored. Use --prop {kv}",
                        Code = "missing_prop_flag",
                        Suggestion = $"--prop {kv}"
                    }).ToList();
                    Console.Error.WriteLine("WARNING: Properties specified without --prop flag.");
                }
                else
                {
                    foreach (var kv in unmatchedKvWarnings)
                        Console.Error.WriteLine($"WARNING: Bare property '{kv}' ignored. Did you mean: --prop {kv}");
                    Console.Error.WriteLine("Hint: Properties must be passed with --prop flag, e.g. officecli add <file> <parent> --type picture --prop src=image.png");
                }
            }

            if (string.IsNullOrEmpty(type) && string.IsNullOrEmpty(from))
            {
                throw new OfficeCli.Core.CliException("Either --type or --from must be specified.")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --type to specify element type, or --from to copy an existing element.",
                    Help = "officecli add <file> <parent> --type <type> --prop src=<file>"
                };
            }

            if (!string.IsNullOrEmpty(from))
            {
                // Copy from existing element
                if (TryResident(file.FullName, req =>
                {
                    req.Command = "add";
                    req.Args["parent"] = parentPath;
                    req.Args["from"] = from;
                    if (position?.Index.HasValue == true) req.Args["index"] = position.Index.Value.ToString();
                    if (position?.After != null) req.Args["after"] = position.After;
                    if (position?.Before != null) req.Args["before"] = position.Before;
                }, json) is {} rc) return rc != 0 ? rc : (hadWarnings ? 2 : 0);

                using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
                var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
                var resultPath = handler.CopyFrom(from, parentPath, position);
                var message = $"Copied to {resultPath}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
                else Console.WriteLine(message);
                if (parentPath == "/") NotifyWatchRoot(handler, file.FullName, oldCount);
                else NotifyWatch(handler, file.FullName, parentPath);
            }
            else
            {
                if (TryResident(file.FullName, req =>
                {
                    req.Command = "add";
                    req.Args["parent"] = parentPath;
                    req.Args["type"] = type!;
                    if (position?.Index.HasValue == true) req.Args["index"] = position.Index.Value.ToString();
                    if (position?.After != null) req.Args["after"] = position.After;
                    if (position?.Before != null) req.Args["before"] = position.Before;
                    req.Props = ParsePropsArray(props);
                }, json) is {} rc) return rc != 0 ? rc : (hadWarnings ? 2 : 0);

                // CONSISTENCY(prop-key-case): --prop keys are case-insensitive
                // so "SRC=x" and "src=x" both resolve to the same handler key.
                // Reuse ParsePropsArray so the inline and resident-server paths
                // stay in sync.
                var properties = ParsePropsArray(props);

                using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
                var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
                var resultPath = handler.Add(parentPath, type!, position, properties);
                var message = $"Added {type!.ToLowerInvariant()} at {resultPath}";
                var spatialLine = GetPptSpatialLine(handler, resultPath);
                var overlapNames = spatialLine != null ? CheckPositionOverlap(handler, resultPath) : new();
                var addWarnings = new List<OfficeCli.Core.CliWarning>();
                if (overlapNames.Count > 0)
                {
                    addWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"Same position as {string.Join(", ", overlapNames)}",
                        Code = "position_overlap",
                        Suggestion = "Use --prop x=... y=... to set distinct positions"
                    });
                }
                var addOverflow = CheckTextOverflow(handler, resultPath);
                if (addOverflow != null)
                {
                    addWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = addOverflow,
                        Code = "text_overflow",
                        Suggestion = "Increase shape height/width, reduce font size, or shorten text"
                    });
                }
                if (json)
                {
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(
                        spatialLine != null ? $"{message}\n  {spatialLine}" : message,
                        addWarnings.Count > 0 ? addWarnings : null));
                }
                else
                {
                    Console.WriteLine(message);
                    if (spatialLine != null) Console.WriteLine($"  {spatialLine}");
                    foreach (var w in addWarnings)
                        Console.Error.WriteLine($"  WARNING: {w.Message}");
                }
                if (parentPath == "/") NotifyWatchRoot(handler, file.FullName, oldCount);
                else NotifyWatch(handler, file.FullName, parentPath);
            }

            return hadWarnings ? 2 : 0;
        }, json); });

        return addCommand;
    }

    private static Command BuildRemoveCommand(Option<bool> jsonOption)
    {
        var removeFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var removePathArg = new Argument<string>("path") { Description = "DOM path of the element to remove" };

        var removeCommand = new Command("remove", "Remove an element from the document");
        removeCommand.Add(removeFileArg);
        removeCommand.Add(removePathArg);
        removeCommand.Add(jsonOption);

        removeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(removeFileArg)!;
            var path = result.GetValue(removePathArg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "remove";
                req.Args["path"] = path;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
            var warning = handler.Remove(path);
            var message = $"Removed {path}";
            if (warning != null) message += $"\n{warning}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            var slideNum = WatchMessage.ExtractSlideNum(path);
            if (slideNum > 0 && !path.Contains("/shape["))
                NotifyWatchRoot(handler, file.FullName, oldCount);
            else
                NotifyWatch(handler, file.FullName, path);
            return 0;
        }, json); });

        return removeCommand;
    }

    private static Command BuildMoveCommand(Option<bool> jsonOption)
    {
        var moveFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var movePathArg = new Argument<string>("path") { Description = "DOM path of the element to move" };
        var moveToOpt = new Option<string?>("--to") { Description = "Target parent path. If omitted, reorders within the current parent" };
        var moveIndexOpt = new Option<int?>("--index") { Description = "Insert position (0-based). If omitted, appends to end" };
        var moveAfterOpt = new Option<string?>("--after") { Description = "Move after the element at this path" };
        var moveBeforeOpt = new Option<string?>("--before") { Description = "Move before the element at this path" };

        var moveCommand = new Command("move", "Move an element to a new position or parent");
        moveCommand.Add(moveFileArg);
        moveCommand.Add(movePathArg);
        moveCommand.Add(moveToOpt);
        moveCommand.Add(moveIndexOpt);
        moveCommand.Add(moveAfterOpt);
        moveCommand.Add(moveBeforeOpt);
        moveCommand.Add(jsonOption);

        moveCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(moveFileArg)!;
            var path = result.GetValue(movePathArg)!;
            var to = result.GetValue(moveToOpt);
            var index = result.GetValue(moveIndexOpt);
            var after = result.GetValue(moveAfterOpt);
            var before = result.GetValue(moveBeforeOpt);

            // Validate mutual exclusivity of --index, --after, --before
            var posCount = (index.HasValue ? 1 : 0) + (after != null ? 1 : 0) + (before != null ? 1 : 0);
            if (posCount > 1)
                throw new OfficeCli.Core.CliException("--index, --after, and --before are mutually exclusive. Use only one.")
                {
                    Code = "invalid_argument",
                    Suggestion = "Use --index for positional insert, or --after/--before for anchor-based insert."
                };

            InsertPosition? position = index.HasValue ? InsertPosition.AtIndex(index.Value)
                : after != null ? InsertPosition.AfterElement(after)
                : before != null ? InsertPosition.BeforeElement(before)
                : null;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "move";
                req.Args["path"] = path;
                if (to != null) req.Args["to"] = to;
                if (position?.Index.HasValue == true) req.Args["index"] = position.Index.Value.ToString();
                if (position?.After != null) req.Args["after"] = position.After;
                if (position?.Before != null) req.Args["before"] = position.Before;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var resultPath = handler.Move(path, to, position);
            var message = $"Moved to {resultPath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            NotifyWatch(handler, file.FullName, path);
            return 0;
        }, json); });

        return moveCommand;
    }

    private static Command BuildSwapCommand(Option<bool> jsonOption)
    {
        var swapFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var swapPath1Arg = new Argument<string>("path1") { Description = "DOM path of the first element" };
        var swapPath2Arg = new Argument<string>("path2") { Description = "DOM path of the second element" };

        var swapCommand = new Command("swap", "Swap two elements in the document");
        swapCommand.Add(swapFileArg);
        swapCommand.Add(swapPath1Arg);
        swapCommand.Add(swapPath2Arg);
        swapCommand.Add(jsonOption);

        swapCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(swapFileArg)!;
            var path1 = result.GetValue(swapPath1Arg)!;
            var path2 = result.GetValue(swapPath2Arg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "swap";
                req.Args["path"] = path1;
                req.Args["to"] = path2;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var (p1, p2) = handler switch
            {
                OfficeCli.Handlers.PowerPointHandler ppt => ppt.Swap(path1, path2),
                OfficeCli.Handlers.WordHandler word => word.Swap(path1, path2),
                OfficeCli.Handlers.ExcelHandler excel => excel.Swap(path1, path2),
                _ => throw new InvalidOperationException("swap not supported for this document type")
            };
            var message = $"Swapped {p1} <-> {p2}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            NotifyWatch(handler, file.FullName, path1);
            return 0;
        }, json); });

        return swapCommand;
    }
}
