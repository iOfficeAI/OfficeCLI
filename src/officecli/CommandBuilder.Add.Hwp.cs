// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static bool IsHwpTextAdd(string extension, string parentPath, string? type)
        => string.Equals(extension, ".hwp", StringComparison.OrdinalIgnoreCase)
            && string.Equals(parentPath, "/text", StringComparison.OrdinalIgnoreCase)
            && (string.IsNullOrWhiteSpace(type)
                || string.Equals(type, "text", StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase));

    private static int HandleHwpTextAdd(
        string inputPath,
        HwpFormat format,
        Dictionary<string, string> properties,
        bool json)
    {
        var value = FirstValue(properties, "value", "text", "content");
        var output = FirstValue(properties, "output", "out");
        if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(output))
        {
            var message = "HWP text add requires --prop value=<text> --prop output=<path>.";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(message));
            else Console.Error.WriteLine(message);
            return 1;
        }

        if (!TryReadInt(properties, 0, out var section, "section", "sec")
            || !TryReadInt(properties, 0, out var paragraph, "paragraph", "para", "p")
            || !TryReadInt(properties, 0, out var offset, "offset", "off"))
        {
            var message = "HWP text add position props must be integers: section, paragraph, offset.";
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
        var engine = HwpEngineSelector.GetEngine(formatKey, HwpCapabilityConstants.OperationInsertText);
        var request = new HwpInsertTextRequest(
            format,
            inputPath,
            outputPath,
            section,
            paragraph,
            offset,
            value,
            json);
        var result = engine.InsertTextAsync(request, CancellationToken.None).GetAwaiter().GetResult();

        if (json)
        {
            var envelope = new System.Text.Json.Nodes.JsonObject
            {
                ["success"] = true,
                ["message"] = $"Inserted HWP text -> {result.OutputPath}",
                ["data"] = new System.Text.Json.Nodes.JsonObject
                {
                    ["outputPath"] = result.OutputPath,
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["evidence"] = HwpCapabilityJsonMapper.ToJsonArray(result.Evidence),
                    ["transaction"] = result.Transaction?.DeepClone()
                },
                ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
            };
            Console.WriteLine(envelope.ToJsonString(OutputFormatter.PublicJsonOptions));
        }
        else
        {
            Console.WriteLine($"Inserted HWP text -> {result.OutputPath}");
            foreach (var warning in result.Warnings)
                Console.Error.WriteLine($"WARNING: {warning}");
        }
        return 0;
    }

    private static bool TryReadInt(
        Dictionary<string, string> properties,
        int defaultValue,
        out int value,
        params string[] keys)
    {
        var raw = FirstValue(properties, keys);
        if (string.IsNullOrWhiteSpace(raw))
        {
            value = defaultValue;
            return true;
        }

        return int.TryParse(raw, out value);
    }
}
