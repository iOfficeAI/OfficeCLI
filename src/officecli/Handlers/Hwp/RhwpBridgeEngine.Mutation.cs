// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp.SafeSave;

namespace OfficeCli.Handlers.Hwp;

public sealed partial class RhwpBridgeEngine
{
    public async Task<HwpMutationResult> FillFieldAsync(HwpFillFieldRequest request, CancellationToken ct)
    {
        if (request.Fields.Count == 0 && request.FieldIds.Count == 0)
            throw new HwpEngineException(
                "fill_field requires at least one field.",
                HwpCapabilityConstants.ReasonUnsupportedOperation,
                "Pass one or more field name/value pairs.",
                [HwpCapabilityConstants.OperationFillField],
                FormatKey(request.Format),
                HwpCapabilityConstants.OperationFillField,
                HwpCapabilityConstants.EngineRhwpBridge,
                HwpCapabilityConstants.ModeExperimental);

        var formatArg = FormatKey(request.Format);
        var currentInput = request.InputPath;
        var tempFiles = new List<string>();
        string? lastOutputJson = null;

        try
        {
            var index = 0;
            foreach (var field in request.Fields)
            {
                index++;
                var output = index == request.Fields.Count + request.FieldIds.Count
                    ? request.OutputPath
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"officecli-rhwp-field-{Guid.NewGuid():N}{Path.GetExtension(request.OutputPath)}");
                if (index != request.Fields.Count) tempFiles.Add(output);

                var args = new[]
                {
                    "set-field", "--format", formatArg, "--input", currentInput,
                    "--output", output, "--name", field.Key, "--value", field.Value, "--json"
                };
                lastOutputJson = await RunBridgeAsync(args, RenderSvgTimeoutMs, ct);
                currentInput = output;
            }
            foreach (var field in request.FieldIds)
            {
                index++;
                var output = index == request.Fields.Count + request.FieldIds.Count
                    ? request.OutputPath
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"officecli-rhwp-field-{Guid.NewGuid():N}{Path.GetExtension(request.OutputPath)}");
                if (index != request.Fields.Count + request.FieldIds.Count) tempFiles.Add(output);

                var args = new[]
                {
                    "set-field", "--format", formatArg, "--input", currentInput,
                    "--output", output, "--id", field.Key.ToString(), "--value", field.Value, "--json"
                };
                lastOutputJson = await RunBridgeAsync(args, RenderSvgTimeoutMs, ct);
                currentInput = output;
            }
        }
        finally
        {
            foreach (var tempFile in tempFiles)
                try { File.Delete(tempFile); } catch { /* best effort */ }
        }

        EnsureOutputExists(request.OutputPath);
        return ParseMutationResult(
            lastOutputJson ?? "{}",
            request.OutputPath,
            "set-field",
            "rhwp-api set-field output file created; caller must verify round-trip before production use.");
    }

    public async Task<HwpMutationResult> ReplaceTextAsync(HwpReplaceTextRequest request, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(request.Query))
            throw new HwpEngineException(
                "replace_text requires a non-empty query.",
                HwpCapabilityConstants.ReasonUnsupportedOperation,
                "Pass --prop find=<text>.",
                [HwpCapabilityConstants.OperationReplaceText],
                FormatKey(request.Format),
                HwpCapabilityConstants.OperationReplaceText,
                HwpCapabilityConstants.EngineRhwpBridge,
                HwpCapabilityConstants.ModeExperimental);

        var formatArg = FormatKey(request.Format);
        string? outputJson = null;
        var runner = new SafeSaveRunner();
        var transaction = await runner.RunAsync(
            new SafeSaveOptions(
                request.InputPath,
                request.OutputPath,
                request.InPlace,
                request.Backup,
                request.Verify,
                HwpCapabilityConstants.OperationReplaceText,
                formatArg,
                SafeSavePolicy.OutputMode("temp-write")),
            async tempPath =>
            {
                var args = new List<string>
                {
                    "replace-text", "--format", formatArg, "--input", request.InputPath,
                    "--output", tempPath, "--query", request.Query,
                    "--value", request.Value, "--mode", request.Mode, "--json"
                };
                if (request.CaseSensitive)
                {
                    args.Add("--case-sensitive");
                    args.Add("true");
                }

                outputJson = await RunBridgeAsync(args.ToArray(), RenderSvgTimeoutMs, ct).ConfigureAwait(false);
                EnsureOutputExists(tempPath);
            },
            _ => Task.FromResult<IReadOnlyList<SafeSaveCheck>>([]),
            ct).ConfigureAwait(false);

        if (!transaction.Ok)
        {
            throw new HwpEngineException(
                "replace_text safe-save transaction failed.",
                HwpCapabilityConstants.ReasonFixtureValidationFailed,
                "Inspect transaction checks and retry with a separate output path.",
                [HwpCapabilityConstants.OperationReplaceText],
                FormatKey(request.Format),
                HwpCapabilityConstants.OperationReplaceText,
                HwpCapabilityConstants.EngineRhwpBridge,
                HwpCapabilityConstants.ModeExperimental,
                SafeSaveJsonMapper.ToJson(transaction));
        }

        EnsureOutputExists(request.OutputPath);
        var result = ParseMutationResult(
            outputJson ?? "{}",
            request.OutputPath,
            "replace-text",
            "rhwp-api replace-text output file created; caller must verify round-trip before production use.");
        return result with { Transaction = SafeSaveJsonMapper.ToJson(transaction) };
    }

    public async Task<HwpMutationResult> SetTableCellAsync(HwpTableCellSetRequest request, CancellationToken ct)
    {
        if (request.Format != HwpFormat.Hwp)
            throw MutationUnsupported(request.Format, HwpCapabilityConstants.OperationSetTableCell);

        var args = new List<string>
        {
            "set-cell-text", "--format", HwpCapabilityConstants.FormatHwp,
            "--input", request.InputPath, "--output", request.OutputPath,
            "--section", request.Section.ToString(),
            "--parent-para", request.ParentParagraph.ToString(),
            "--control", request.Control.ToString(),
            "--cell", request.Cell.ToString(),
            "--cell-para", request.CellParagraph.ToString(),
            "--offset", request.Offset.ToString(),
            "--value", request.Value,
            "--json"
        };
        if (request.Count.HasValue)
        {
            args.Add("--count");
            args.Add(request.Count.Value.ToString());
        }

        var outputJson = await RunBridgeAsync(args.ToArray(), RenderSvgTimeoutMs, ct);
        EnsureOutputExists(request.OutputPath);
        return ParseMutationResult(
            outputJson,
            request.OutputPath,
            "set-cell-text",
            "rhwp-api set-cell-text output file created; verified only for HWP input fixtures.");
    }

    public Task<HwpMutationResult> SaveOriginalAsync(HwpSaveOriginalRequest request, CancellationToken ct)
        => Task.FromException<HwpMutationResult>(
            MutationUnsupported(request.Format, HwpCapabilityConstants.OperationSaveOriginal));

    public Task<HwpMutationResult> SaveAsHwpAsync(HwpSaveAsHwpRequest request, CancellationToken ct)
        => Task.FromException<HwpMutationResult>(
            MutationUnsupported(request.Format, HwpCapabilityConstants.OperationSaveAsHwp));

    private static void EnsureOutputExists(string outputPath)
    {
        if (File.Exists(outputPath)) return;
        throw new HwpEngineException(
            "rhwp-officecli-bridge did not create the requested output file.",
            HwpCapabilityConstants.ReasonBridgeExitNonZero,
            engine: HwpCapabilityConstants.EngineRhwpBridge,
            engineMode: HwpCapabilityConstants.ModeExperimental);
    }

    private static HwpMutationResult ParseMutationResult(
        string json,
        string outputPath,
        string bridgeCommand,
        string evidence)
    {
        JsonNode? node;
        try { node = JsonNode.Parse(json); }
        catch
        {
            throw new HwpEngineException(
                $"rhwp-officecli-bridge {bridgeCommand} produced unparseable JSON.",
                HwpCapabilityConstants.ReasonBridgeInvalidJson,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);
        }

        var engineVersion = node?["engineVersion"]?.GetValue<string>();
        var warnings = ParseWarnings(node);
        return new HwpMutationResult(
            outputPath,
            HwpCapabilityConstants.EngineRhwpBridge,
            engineVersion,
            [evidence],
            warnings);
    }

    private static string FormatKey(HwpFormat format)
        => format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;

    private static HwpEngineException MutationUnsupported(HwpFormat format, string operation)
    {
        var formatKey = FormatKey(format);
        var reason = operation switch
        {
            HwpCapabilityConstants.OperationFillField when format == HwpFormat.Hwp
                => HwpCapabilityConstants.ReasonBinaryHwpMutationForbidden,
            HwpCapabilityConstants.OperationSaveOriginal or HwpCapabilityConstants.OperationSaveAsHwp when format == HwpFormat.Hwp
                => HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden,
            _ => HwpCapabilityConstants.ReasonRoundTripUnverified
        };
        return new HwpEngineException(
            $"{formatKey} operation '{operation}' is not supported in Phase 0.5 (read/render only).",
            reason,
            "Use explicit output-file mutation commands only in experimental bridge mode.",
            [HwpCapabilityConstants.OperationReadText, HwpCapabilityConstants.OperationRenderSvg],
            formatKey,
            operation,
            HwpCapabilityConstants.EngineRhwpBridge,
            HwpCapabilityConstants.ModeExperimental);
    }
}
