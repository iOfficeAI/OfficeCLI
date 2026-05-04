// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Security.Cryptography;
using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwpx.Validation;
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
        var sourceHash = ComputeSha256(request.InputPath);
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
                BuildReplaceTextSafeSavePolicy(request.Format)),
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
            tempPath => ValidateReplaceTextAsync(request, tempPath, sourceHash, ct),
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

    private async Task<SafeSaveValidationResult> ValidateReplaceTextAsync(
        HwpReplaceTextRequest request,
        string tempPath,
        string sourceHash,
        CancellationToken ct)
    {
        var checks = new List<SafeSaveCheck>();
        var semanticDelta = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["query"] = request.Query,
            ["replacement"] = request.Value,
            ["mode"] = request.Mode
        };

        try
        {
            var sourceText = await ReadTextOnlyAsync(request.Format, request.InputPath, ct).ConfigureAwait(false);
            var outputText = await ReadTextOnlyAsync(request.Format, tempPath, ct).ConfigureAwait(false);
            checks.Add(new SafeSaveCheck(
                "provider-readback",
                true,
                "info",
                null,
                new Dictionary<string, object?>
                {
                    ["sourceLength"] = sourceText.Length,
                    ["outputLength"] = outputText.Length
                }));

            var comparison = request.CaseSensitive
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;
            var sourceOldCount = CountOccurrences(sourceText, request.Query, comparison);
            var outputOldCount = CountOccurrences(outputText, request.Query, comparison);
            var outputNewCount = string.IsNullOrEmpty(request.Value)
                ? 0
                : CountOccurrences(outputText, request.Value, comparison);
            var valueContainsQuery = !string.IsNullOrEmpty(request.Value)
                && request.Value.Contains(request.Query, comparison);
            var allMode = string.Equals(request.Mode, "all", StringComparison.OrdinalIgnoreCase);
            var oldTextChanged = valueContainsQuery
                ? !string.Equals(sourceText, outputText, StringComparison.Ordinal)
                : allMode
                    ? outputOldCount == 0
                    : outputOldCount == Math.Max(0, sourceOldCount - 1);
            var newTextPresent = string.IsNullOrEmpty(request.Value) || outputNewCount > 0;
            var semanticOk = sourceOldCount > 0
                && oldTextChanged
                && newTextPresent
                && !string.Equals(sourceText, outputText, StringComparison.Ordinal);

            semanticDelta["sourceOldCount"] = sourceOldCount;
            semanticDelta["outputOldCount"] = outputOldCount;
            semanticDelta["outputNewCount"] = outputNewCount;
            semanticDelta["changed"] = semanticOk;

            checks.Add(new SafeSaveCheck(
                "semantic-delta",
                semanticOk,
                semanticOk ? "info" : "error",
                semanticOk ? null : "Replacement output did not contain the expected semantic delta.",
                semanticDelta));
        }
        catch (Exception ex)
        {
            checks.Add(new SafeSaveCheck(
                "provider-readback",
                false,
                "error",
                ex.Message));
        }

        var currentSourceHash = ComputeSha256(request.InputPath);
        var sourcePreserved = string.Equals(sourceHash, currentSourceHash, StringComparison.OrdinalIgnoreCase);
        checks.Add(new SafeSaveCheck(
            "source-preserved",
            sourcePreserved,
            sourcePreserved ? "info" : "error",
            sourcePreserved ? null : "Source file changed before safe-save commit.",
            new Dictionary<string, object?>
            {
                ["beforeSha256"] = sourceHash,
                ["afterSha256"] = currentSourceHash
            }));

        Dictionary<string, object?>? packageIntegrity = null;
        if (request.Format == HwpFormat.Hwpx)
        {
            var packageResult = HwpxPackageValidator.Validate(tempPath);
            checks.AddRange(packageResult.Checks);
            packageIntegrity = new Dictionary<string, object?>(packageResult.PackageIntegrity, StringComparer.Ordinal);
        }

        IReadOnlyDictionary<string, object?>? visualDelta = null;
        if (request.Verify)
        {
            var visualResult = await ValidateVisualAsync(request.Format, tempPath, ct).ConfigureAwait(false);
            checks.AddRange(visualResult.Checks);
            visualDelta = visualResult.VisualDelta;
        }

        return new SafeSaveValidationResult(checks, semanticDelta, visualDelta, packageIntegrity);
    }

    private static SafeSavePolicy BuildReplaceTextSafeSavePolicy(HwpFormat format)
    {
        var required = new List<string>
        {
            "temp-write",
            "provider-readback",
            "semantic-delta",
            "source-preserved"
        };
        if (format == HwpFormat.Hwpx)
            required.Add("package-integrity");
        return SafeSavePolicy.OutputMode(required.ToArray());
    }

    private async Task<SafeSaveValidationResult> ValidateVisualAsync(
        HwpFormat format,
        string tempPath,
        CancellationToken ct)
    {
        var outputDir = Path.Combine(
            Path.GetDirectoryName(tempPath) ?? Path.GetTempPath(),
            $".{Path.GetFileNameWithoutExtension(tempPath)}.svg");
        try
        {
            var renderResult = await RenderSvgAsync(
                new HwpRenderRequest(
                    format,
                    tempPath,
                    outputDir,
                    "1",
                    new FileInfo(tempPath).Length,
                    Json: false),
                ct).ConfigureAwait(false);
            return SafeSaveVisualValidator.FromRenderResult(renderResult);
        }
        catch (Exception ex)
        {
            return SafeSaveVisualValidator.FromFailure(ex);
        }
    }

    private async Task<string> ReadTextOnlyAsync(HwpFormat format, string path, CancellationToken ct)
    {
        var result = await ReadTextAsync(
            new HwpReadRequest(format, path, new FileInfo(path).Length, Json: true),
            ct).ConfigureAwait(false);
        return result.Text;
    }

    private static int CountOccurrences(string text, string value, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value)) return 0;
        var count = 0;
        var index = 0;
        while (index < text.Length)
        {
            var found = text.IndexOf(value, index, comparison);
            if (found < 0) break;
            count++;
            index = found + value.Length;
        }
        return count;
    }

    private static string ComputeSha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
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
