// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
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
        bool json,
        bool inPlace,
        bool backup,
        bool verify)
    {
        var query = FirstValue(properties, "find", "query", "old");
        var value = FirstValue(properties, "value", "text", "new");
        var output = FirstValue(properties, "output", "out");
        var mode = FirstValue(properties, "mode") ?? "one";
        var caseSensitiveRaw = FirstValue(properties, "case-sensitive", "caseSensitive") ?? "false";
        if (string.IsNullOrEmpty(query) || value == null || (!inPlace && string.IsNullOrWhiteSpace(output)))
        {
            var message = "HWP text replace requires --prop find=<text> --prop value=<text> plus --prop output=<path> or --in-place.";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(message));
            else Console.Error.WriteLine(message);
            return 1;
        }
        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
        if (inPlace && !string.IsNullOrWhiteSpace(output))
            return WriteHwpTextReplacePolicyError(
                json,
                formatKey,
                "hwp_in_place_output_conflict",
                "Use either --in-place or --prop output=<path>, not both.",
                "Remove --prop output or remove --in-place.",
                "officecli help hwp");
        if (inPlace && !backup)
            return WriteHwpTextReplacePolicyError(
                json,
                formatKey,
                "hwp_in_place_requires_backup",
                "HWP in-place text replacement requires --backup.",
                "Add --backup or use --prop output=<path> instead.",
                "officecli help hwp");
        if (inPlace && !verify)
            return WriteHwpTextReplacePolicyError(
                json,
                formatKey,
                "hwp_in_place_requires_verify",
                "HWP in-place text replacement requires --verify.",
                "Add --verify or use --prop output=<path> instead.",
                "officecli help hwp");

        var outputPath = inPlace ? Path.GetFullPath(inputPath) : Path.GetFullPath(output!);
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir)) Directory.CreateDirectory(outputDir);

        var engine = HwpEngineSelector.GetEngine(formatKey, HwpCapabilityConstants.OperationReplaceText);
        var request = new HwpReplaceTextRequest(
            format,
            inputPath,
            outputPath,
            query,
            value,
            mode,
            bool.TryParse(caseSensitiveRaw, out var caseSensitive) && caseSensitive,
            inPlace,
            backup,
            verify,
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
                    ["evidence"] = HwpCapabilityJsonMapper.ToJsonArray(result.Evidence),
                    ["transaction"] = result.Transaction?.DeepClone()
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

    private static int WriteHwpTextReplacePolicyError(
        bool json,
        string format,
        string code,
        string message,
        string suggestion,
        string nextCommand)
    {
        if (json)
        {
            var envelope = new System.Text.Json.Nodes.JsonObject
            {
                ["success"] = false,
                ["error"] = new System.Text.Json.Nodes.JsonObject
                {
                    ["error"] = message,
                    ["code"] = code,
                    ["suggestion"] = suggestion,
                    ["help"] = "officecli help hwp",
                    ["format"] = format,
                    ["operation"] = HwpCapabilityConstants.OperationReplaceText,
                    ["engine"] = HwpCapabilityConstants.EngineRhwpBridge,
                    ["engineMode"] = HwpCapabilityConstants.ModeExperimental,
                    ["nextCommand"] = nextCommand
                }
            };
            Console.WriteLine(envelope.ToJsonString(OutputFormatter.PublicJsonOptions));
        }
        else
        {
            Console.Error.WriteLine(message);
            Console.Error.WriteLine($"Hint: {suggestion}");
        }
        return 1;
    }
}
