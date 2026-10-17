// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildRawCommand(Option<bool> jsonOption)
    {
        var rawFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var rawPathArg = new Argument<string>("part") { Description = "Part path (e.g. /document, /styles, /header[1])" };
        rawPathArg.DefaultValueFactory = _ => "/document";

        var rawStartOpt = new Option<int?>("--start") { Description = "Start row number (Excel sheets only)" };
        var rawEndOpt = new Option<int?>("--end") { Description = "End row number (Excel sheets only)" };

        var rawColsOpt = new Option<string?>("--cols") { Description = "Column filter, comma-separated (Excel only, e.g. A,B,C)" };

        var rawCommand = new Command("raw", "View raw XML of a document part");
        rawCommand.Add(rawFileArg);
        rawCommand.Add(rawPathArg);
        rawCommand.Add(rawStartOpt);
        rawCommand.Add(rawEndOpt);
        rawCommand.Add(rawColsOpt);
        rawCommand.Add(jsonOption);

        rawCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(rawFileArg)!;
            var partPath = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(rawPathArg)!)!;
            var startRow = result.GetValue(rawStartOpt);
            var endRow = result.GetValue(rawEndOpt);
            var rawColsStr = result.GetValue(rawColsOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "raw";
                req.Args["part"] = partPath;
                if (startRow.HasValue) req.Args["start"] = startRow.Value.ToString();
                if (endRow.HasValue) req.Args["end"] = endRow.Value.ToString();
                if (rawColsStr != null) req.Args["cols"] = rawColsStr;
            }, json) is {} rc) return rc;

            var rawCols = rawColsStr != null ? new HashSet<string>(rawColsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var xml = handler.Raw(partPath, startRow, endRow, rawCols);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(xml));
            else Console.WriteLine(xml);
            return 0;
        }, json); });

        return rawCommand;
    }

    private static Command BuildRawSetCommand(Option<bool> jsonOption)
    {
        var rawSetFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var rawSetPartArg = new Argument<string>("part") { Description = "Part path (e.g. /document, /styles, /Sheet1, /slide[1])" };
        var rawSetXpathOpt = new Option<string>("--xpath") { Description = "XPath to target element(s)", Required = true };
        var rawSetActionOpt = new Option<string>("--action") { Description = "Action: append, prepend, insertbefore, insertafter, replace, remove, setattr", Required = true };
        var rawSetXmlOpt = new Option<string?>("--xml") { Description = "XML fragment or attr=value for setattr" };

        var rawSetCommand = new Command("raw-set", "Modify raw XML in a document part (universal fallback for any OpenXML operation)");
        rawSetCommand.Add(rawSetFileArg);
        rawSetCommand.Add(rawSetPartArg);
        rawSetCommand.Add(rawSetXpathOpt);
        rawSetCommand.Add(rawSetActionOpt);
        rawSetCommand.Add(rawSetXmlOpt);
        rawSetCommand.Add(jsonOption);

        rawSetCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(rawSetFileArg)!;
            var partPath = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(rawSetPartArg)!)!;
            var xpath = result.GetValue(rawSetXpathOpt)!;
            var action = result.GetValue(rawSetActionOpt)!;
            var xml = result.GetValue(rawSetXmlOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "raw-set";
                req.Args["part"] = partPath;
                req.Args["xpath"] = xpath;
                req.Args["action"] = action;
                if (xml != null) req.Args["xml"] = xml;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var errorsBefore = handler.Validate().Select(e => e.Description).ToHashSet();
            handler.RawSet(partPath, xpath, action, xml);
            var warnings = ReportNewErrorsAsWarnings(handler, errorsBefore);
            var message = $"raw-set applied: {action} at {xpath}";
            handler.Persist(); // issue #517: write before reporting — see IDocumentHandler.Persist
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message, warnings));
            else
            {
                Console.WriteLine(message);
                ReportNewErrors(handler, errorsBefore, warnings);
            }
            NotifyWatch(handler, file.FullName, null);
            // The edit IS applied when the SDK validator gains new errors — the
            // validator is advisory (it flags element order Word itself accepts),
            // so raw-set stays the escape hatch and never rolls back. Report
            // "applied with caveats" as exit 2, the same code add/set use for
            // unsupported_property. Exit 1 read as "failed, retry", and a retry
            // appended the content a second time (issue #374).
            return warnings is { Count: > 0 } ? 2 : 0;
        }, json); });

        return rawSetCommand;
    }

    private static Command BuildAddPartCommand(Option<bool> jsonOption)
    {
        var addPartFileArg = new Argument<string>("file") { Description = "Document file path" };
        var addPartParentArg = new Argument<string>("parent") { Description = "Parent part path (e.g. / for document root, /Sheet1 for Excel sheet, /slide[0] for PPT slide)" };
        var addPartTypeOpt = new Option<string>("--type") { Description = "Part type to create. Word: chart, header, footer. PPT/Excel: chart", Required = true };
        var addPartCommand = new Command("add-part", "Create a new document part and return its relationship ID for use with raw-set");
        addPartCommand.Add(addPartFileArg);
        addPartCommand.Add(addPartParentArg);
        addPartCommand.Add(addPartTypeOpt);
        addPartCommand.Add(jsonOption);

        addPartCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(addPartFileArg)!;
            var parent = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(addPartParentArg)!)!;
            var type = result.GetValue(addPartTypeOpt)!;

            if (TryResident(file, req =>
            {
                req.Command = "add-part";
                req.Args["parent"] = parent;
                req.Args["type"] = type;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file, editable: true);
            var errorsBefore = handler.Validate().Select(e => e.Description).ToHashSet();
            var (relId, partPath) = handler.AddPart(parent, type);
            var warnings = ReportNewErrorsAsWarnings(handler, errorsBefore);
            var message = $"Created {type} part: relId={relId} path={partPath}";
            handler.Persist(); // issue #517: write before reporting — see IDocumentHandler.Persist
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message, warnings));
            else
            {
                Console.WriteLine(message);
                ReportNewErrors(handler, errorsBefore, warnings);
            }
            NotifyWatch(handler, file, null);
            // Same contract as raw-set: applied, validator caveats → exit 2.
            return warnings is { Count: > 0 } ? 2 : 0;
        }, json); });

        return addPartCommand;
    }
}
