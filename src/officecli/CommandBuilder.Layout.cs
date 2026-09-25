// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildLayoutCommand(Option<bool> jsonOption)
    {
        var layoutFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var layoutPathArg = new Argument<string>("path") { Description = "Slide DOM path (e.g. /slide[2]). Geometric layout operates per slide." };
        var layoutAlignOpt = new Option<string?>("--align") { Description = "Align shapes in one call: left | center | right | top | middle | bottom (relative to the targeted shapes' bounding box), or slide-left | slide-center | slide-right | slide-top | slide-middle | slide-bottom (relative to the slide). NOTE: this is geometric; align= on a SHAPE path is text alignment." };
        var layoutDistributeOpt = new Option<string?>("--distribute") { Description = "Distribute shapes evenly in one call: horizontal | vertical (aliases h/horiz, v/vert). Gaps become equal; endpoints stay fixed. Needs at least 3 shapes." };
        var layoutTargetsOpt = new Option<string?>("--targets") { Description = "Comma-separated shape scope: shape[N], bare N, or shape[@id=N] (stable; prefer the @id paths get/query emit). Omit = all auto-shapes/textboxes on the slide. Pictures/charts/tables/connectors are not matched." };

        var layoutCommand = new Command("layout", "Align and/or distribute shapes on a slide in one call (geometric layout; pptx only)");
        layoutCommand.Add(layoutFileArg);
        layoutCommand.Add(layoutPathArg);
        layoutCommand.Add(layoutAlignOpt);
        layoutCommand.Add(layoutDistributeOpt);
        layoutCommand.Add(layoutTargetsOpt);
        layoutCommand.Add(jsonOption);

        layoutCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(layoutFileArg)!;
            var path = MsysPathHint.Restore(result.GetValue(layoutPathArg)!)!;
            var align = result.GetValue(layoutAlignOpt);
            var distribute = result.GetValue(layoutDistributeOpt);
            var targets = MsysPathHint.Restore(result.GetValue(layoutTargetsOpt));

            OfficeCli.Core.MutationSelectorGuard.EnsureScoped(path, "layout");

            if (TryResident(file.FullName, req =>
            {
                req.Command = "layout";
                req.Args["path"] = path;
                if (!string.IsNullOrWhiteSpace(align)) req.Args["align"] = align;
                if (!string.IsNullOrWhiteSpace(distribute)) req.Args["distribute"] = distribute;
                if (!string.IsNullOrWhiteSpace(targets)) req.Args["targets"] = targets;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            if (handler is not OfficeCli.Handlers.PowerPointHandler pptHandler)
                throw new OfficeCli.Core.CliException("'layout' is only supported for .pptx files.")
                    { Code = "unsupported_type" };
            var oldCount = pptHandler.GetSlideCount();
            var message = pptHandler.LayoutSlide(path, align, distribute, targets);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            var slideNum = WatchMessage.ExtractSlideNum(path);
            if (slideNum > 0 && !path.Contains("/shape["))
                NotifyWatchRoot(handler, file.FullName, oldCount);
            else
                NotifyWatch(handler, file.FullName, path);
            return 0;
        }, json); });

        return layoutCommand;
    }
}
