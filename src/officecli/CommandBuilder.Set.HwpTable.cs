// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static int HandleHwpTableCellSet(
        string inputPath,
        HwpFormat format,
        Dictionary<string, string> properties,
        bool json)
    {
        var output = FirstValue(properties, "output", "out");
        var value = FirstValue(properties, "value", "text");
        if (format != HwpFormat.Hwp)
            return HwpTableCellError("HWP table cell set is currently supported only for .hwp input.", json);
        if (string.IsNullOrWhiteSpace(output) || value == null)
            return HwpTableCellError("HWP table cell set requires --prop value=<text> --prop output=<path>.", json);
        if (!TryReadInt(properties, out var section, "section", "sec")
            || !TryReadInt(properties, out var parentPara, "parent-para", "parentParagraph", "paragraph", "para")
            || !TryReadInt(properties, out var control, "control", "control-index", "controlIndex")
            || !TryReadInt(properties, out var cell, "cell", "cell-index", "cellIndex"))
        {
            return HwpTableCellError(
                "HWP table cell set requires numeric section, parent-para, control, and cell props.",
                json);
        }

        var cellPara = TryReadInt(properties, out var parsedCellPara, "cell-para", "cellParagraph", "cell-paragraph")
            ? parsedCellPara
            : 0;
        var offset = TryReadInt(properties, out var parsedOffset, "offset") ? parsedOffset : 0;
        int? count = TryReadInt(properties, out var parsedCount, "count") ? parsedCount : null;
        var outputPath = Path.GetFullPath(output);
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir)) Directory.CreateDirectory(outputDir);

        var engine = HwpEngineSelector.GetEngine(
            HwpCapabilityConstants.FormatHwp,
            HwpCapabilityConstants.OperationSetTableCell);
        var request = new HwpTableCellSetRequest(
            format,
            inputPath,
            outputPath,
            section,
            parentPara,
            control,
            cell,
            cellPara,
            offset,
            count,
            value,
            json);
        var result = engine.SetTableCellAsync(request, CancellationToken.None).GetAwaiter().GetResult();

        if (json)
        {
            var envelope = new System.Text.Json.Nodes.JsonObject
            {
                ["success"] = true,
                ["message"] = $"Updated HWP table cell ({parentPara},{control},{cell}) -> {result.OutputPath}",
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
            Console.WriteLine($"Updated HWP table cell ({parentPara},{control},{cell}) -> {result.OutputPath}");
            foreach (var warning in result.Warnings)
                Console.Error.WriteLine($"WARNING: {warning}");
        }
        return 0;
    }

    private static bool TryReadInt(
        Dictionary<string, string> properties,
        out int value,
        params string[] keys)
    {
        value = 0;
        var raw = FirstValue(properties, keys);
        return raw != null && int.TryParse(raw, out value);
    }

    private static int HwpTableCellError(string message, bool json)
    {
        if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(message));
        else Console.Error.WriteLine(message);
        return 1;
    }
}
