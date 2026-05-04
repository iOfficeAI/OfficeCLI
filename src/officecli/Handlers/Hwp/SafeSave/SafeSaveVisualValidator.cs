// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal static class SafeSaveVisualValidator
{
    public static SafeSaveValidationResult FromRenderResult(HwpRenderResult renderResult)
    {
        var pageCount = renderResult.Pages.Count;
        var firstPage = renderResult.Pages.FirstOrDefault();
        var visualDelta = new Dictionary<string, object?>
        {
            ["pageCount"] = pageCount,
            ["manifestPath"] = renderResult.ManifestPath,
            ["firstPageSha256"] = firstPage?.Sha256,
            ["engine"] = renderResult.Engine,
            ["engineVersion"] = renderResult.EngineVersion
        };
        var check = new SafeSaveCheck(
            "visual-render",
            pageCount > 0,
            pageCount > 0 ? "info" : "warning",
            pageCount > 0 ? null : "Provider SVG render returned no pages.",
            visualDelta);
        return new SafeSaveValidationResult([check], VisualDelta: visualDelta);
    }

    public static SafeSaveValidationResult FromFailure(Exception exception)
    {
        var visualDelta = new Dictionary<string, object?>
        {
            ["error"] = exception.Message
        };
        return new SafeSaveValidationResult(
            [new SafeSaveCheck("visual-render", false, "warning", exception.Message, visualDelta)],
            VisualDelta: visualDelta);
    }
}
