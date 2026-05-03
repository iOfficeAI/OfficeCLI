// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Text.Json.Nodes;

namespace OfficeCli.Handlers.Hwp;

/// <summary>
/// Phase 0.5 experimental engine that routes read/render operations through
/// the rhwp-officecli-bridge subprocess. Only active when
/// OFFICECLI_HWP_ENGINE=rhwp-experimental. Mutation operations are forbidden.
/// </summary>
public sealed class RhwpBridgeEngine : IHwpEngine
{
    private const string BridgeExecutableName = "rhwp-officecli-bridge";
    private const int CapabilitiesTimeoutMs = 2_000;
    private const int ReadTextTimeoutMs = 10_000;
    private const int RenderSvgTimeoutMs = 60_000;
    private const long LargeFileSizeBytes = 10L * 1024 * 1024;

    private readonly string _bridgePath;

    private RhwpBridgeEngine(string bridgePath)
    {
        _bridgePath = bridgePath;
    }

    public string Name => HwpCapabilityConstants.EngineRhwpBridge;
    public string? Version => null;
    public HwpEngineMode Mode => HwpEngineMode.Experimental;

    /// <summary>
    /// Returns a new bridge engine if the bridge executable can be located,
    /// or null with a reason description.
    /// </summary>
    public static RhwpBridgeEngine? TryCreate(out string? missingReason)
    {
        missingReason = null;
        var bridgePath = DiscoverBridge();
        if (bridgePath == null)
        {
            missingReason = "bridge not found in OFFICECLI_RHWP_BRIDGE_PATH, executable directory, or PATH";
            return null;
        }
        return new RhwpBridgeEngine(bridgePath);
    }

    public Task<HwpCapabilityReport> GetCapabilitiesAsync(CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        return Task.FromResult(HwpCapabilityFactory.BuildReport(Name));
    }

    public async Task<HwpTextResult> ReadTextAsync(HwpReadRequest request, CancellationToken ct)
    {
        var formatArg = request.Format == HwpFormat.Hwp ? "hwp" : "hwpx";
        var timeout = request.InputSizeBytes > LargeFileSizeBytes
            ? ReadTextTimeoutMs * 3
            : ReadTextTimeoutMs;
        var args = new[] { "read-text", "--format", formatArg, "--input", request.InputPath, "--json" };
        var output = await RunBridgeAsync(args, timeout, ct);
        return ParseTextResult(output);
    }

    public async Task<HwpRenderResult> RenderSvgAsync(HwpRenderRequest request, CancellationToken ct)
    {
        var formatArg = request.Format == HwpFormat.Hwp ? "hwp" : "hwpx";
        var timeout = request.InputSizeBytes > LargeFileSizeBytes
            ? RenderSvgTimeoutMs * 3
            : RenderSvgTimeoutMs;
        var args = new[]
        {
            "render-svg", "--format", formatArg, "--input", request.InputPath,
            "--out-dir", request.OutputDirectory, "--page", request.PageSelector, "--json"
        };
        var output = await RunBridgeAsync(args, timeout, ct);
        return ParseRenderResult(output, request.OutputDirectory);
    }

    public Task<HwpMutationResult> FillFieldAsync(HwpFillFieldRequest request, CancellationToken ct)
        => Task.FromException<HwpMutationResult>(
            MutationUnsupported(request.Format, HwpCapabilityConstants.OperationFillField));

    public Task<HwpMutationResult> SaveOriginalAsync(HwpSaveOriginalRequest request, CancellationToken ct)
        => Task.FromException<HwpMutationResult>(
            MutationUnsupported(request.Format, HwpCapabilityConstants.OperationSaveOriginal));

    public Task<HwpMutationResult> SaveAsHwpAsync(HwpSaveAsHwpRequest request, CancellationToken ct)
        => Task.FromException<HwpMutationResult>(
            MutationUnsupported(request.Format, HwpCapabilityConstants.OperationSaveAsHwp));

    private async Task<string> RunBridgeAsync(string[] args, int timeoutMs, CancellationToken ct)
    {
        if (!File.Exists(_bridgePath))
            throw new HwpEngineException(
                $"rhwp-officecli-bridge not found at '{_bridgePath}'.",
                HwpCapabilityConstants.ReasonBridgeMissing,
                "Set OFFICECLI_RHWP_BRIDGE_PATH or install the bridge beside officecli.",
                [],
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(timeoutMs);

        var psi = new ProcessStartInfo
        {
            FileName = BridgeFileName(_bridgePath),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetExtension(_bridgePath), ".dll", StringComparison.OrdinalIgnoreCase))
            psi.ArgumentList.Add(_bridgePath);
        foreach (var arg in args)
            psi.ArgumentList.Add(arg);

        Process? process;
        try
        {
            process = Process.Start(psi);
        }
        catch (Exception ex)
        {
            throw new HwpEngineException(
                $"Failed to start rhwp-officecli-bridge: {ex.Message}",
                HwpCapabilityConstants.ReasonBridgeMissing,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);
        }

        if (process == null)
            throw new HwpEngineException(
                "Failed to start rhwp-officecli-bridge process.",
                HwpCapabilityConstants.ReasonBridgeMissing,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);

        string stdout;
        int exitCode;
        try
        {
            var stdoutTask = process.StandardOutput.ReadToEndAsync(cts.Token);
            var stderrTask = process.StandardError.ReadToEndAsync(cts.Token);
            await process.WaitForExitAsync(cts.Token);
            stdout = await stdoutTask;
            _ = await stderrTask;
            exitCode = process.ExitCode;
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            try { process.Kill(entireProcessTree: true); } catch { /* best effort */ }
            process.Dispose();
            throw new HwpEngineException(
                $"rhwp-officecli-bridge timed out after {timeoutMs}ms.",
                HwpCapabilityConstants.ReasonBridgeTimeout,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);
        }
        finally
        {
            process.Dispose();
        }

        if (exitCode != 0)
            throw new HwpEngineException(
                $"rhwp-officecli-bridge exited with code {exitCode}.",
                HwpCapabilityConstants.ReasonBridgeExitNonZero,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);

        var trimmed = stdout.Trim();
        if (string.IsNullOrEmpty(trimmed) || trimmed[0] != '{')
            throw new HwpEngineException(
                "rhwp-officecli-bridge produced invalid JSON output.",
                HwpCapabilityConstants.ReasonBridgeInvalidJson,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);

        return trimmed;
    }

    private static HwpTextResult ParseTextResult(string json)
    {
        JsonNode? node;
        try { node = JsonNode.Parse(json); }
        catch
        {
            throw new HwpEngineException(
                "rhwp-officecli-bridge read-text produced unparseable JSON.",
                HwpCapabilityConstants.ReasonBridgeInvalidJson,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);
        }

        var text = node?["text"]?.GetValue<string>() ?? "";
        var engineVersion = node?["engineVersion"]?.GetValue<string>();
        var pages = new List<HwpTextPage>();
        if (node?["pages"] is JsonArray pagesArr)
        {
            foreach (var p in pagesArr)
            {
                var pageNum = p?["page"]?.GetValue<int>() ?? 0;
                var pageText = p?["text"]?.GetValue<string>() ?? "";
                pages.Add(new HwpTextPage(pageNum, pageText));
            }
        }
        var warnings = new List<string>();
        if (node?["warnings"] is JsonArray wArr)
            foreach (var w in wArr)
                if (w?.GetValue<string>() is { } ws) warnings.Add(ws);

        return new HwpTextResult(
            text, pages,
            HwpCapabilityConstants.EngineRhwpBridge,
            engineVersion, [], warnings.ToArray());
    }

    private static HwpRenderResult ParseRenderResult(string json, string outputDir)
    {
        JsonNode? node;
        try { node = JsonNode.Parse(json); }
        catch
        {
            throw new HwpEngineException(
                "rhwp-officecli-bridge render-svg produced unparseable JSON.",
                HwpCapabilityConstants.ReasonBridgeInvalidJson,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);
        }

        var pages = new List<HwpRenderedPage>();
        if (node?["pages"] is JsonArray pagesArr)
        {
            foreach (var p in pagesArr)
            {
                var pageNum = p?["page"]?.GetValue<int>() ?? 0;
                var svgPath = p?["path"]?.GetValue<string>() ?? "";
                var sha256 = p?["sha256"]?.GetValue<string>() ?? "";
                pages.Add(new HwpRenderedPage(pageNum, svgPath, sha256));
            }
        }
        var manifestPath = node?["manifest"]?.GetValue<string>()
            ?? Path.Combine(outputDir, "manifest.json");
        var engineVersion = node?["engineVersion"]?.GetValue<string>();
        var warnings = new List<string>();
        if (node?["warnings"] is JsonArray wArr)
            foreach (var w in wArr)
                if (w?.GetValue<string>() is { } ws) warnings.Add(ws);

        return new HwpRenderResult(
            pages, manifestPath,
            HwpCapabilityConstants.EngineRhwpBridge,
            engineVersion, [], warnings.ToArray());
    }

    private static string? DiscoverBridge()
    {
        var envPath = Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH");
        if (!string.IsNullOrEmpty(envPath) && File.Exists(envPath))
            return envPath;

        var exeDir = Path.GetDirectoryName(
            Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName);
        if (exeDir != null)
        {
            foreach (var name in new[] { BridgeExecutableName, BridgeExecutableName + ".exe", BridgeExecutableName + ".dll" })
            {
                var candidate = Path.Combine(exeDir, name);
                if (File.Exists(candidate)) return candidate;
            }
        }

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            foreach (var name in new[] { BridgeExecutableName, BridgeExecutableName + ".exe", BridgeExecutableName + ".dll" })
            {
                var candidate = Path.Combine(dir, name);
                if (File.Exists(candidate)) return candidate;
            }
        }

        return null;
    }

    private static string BridgeFileName(string bridgePath)
        => string.Equals(Path.GetExtension(bridgePath), ".dll", StringComparison.OrdinalIgnoreCase)
            ? "dotnet"
            : bridgePath;

    private static HwpEngineException MutationUnsupported(HwpFormat format, string operation)
    {
        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
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
            "Use read_text or render_svg only in Phase 0.5.",
            [HwpCapabilityConstants.OperationReadText, HwpCapabilityConstants.OperationRenderSvg],
            formatKey,
            operation,
            HwpCapabilityConstants.EngineRhwpBridge,
            HwpCapabilityConstants.ModeExperimental);
    }
}
