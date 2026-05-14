// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp;

/// <summary>
/// Selects the active HWP engine based on the OFFICECLI_HWP_ENGINE environment variable.
/// Default: CustomHwpxEngine (existing XML-first behavior).
/// Experimental: RhwpBridgeEngine when OFFICECLI_HWP_ENGINE=rhwp-experimental.
/// </summary>
public static class HwpEngineSelector
{
    private const string EnvVarName = "OFFICECLI_HWP_ENGINE";
    private const string ExperimentalValue = "rhwp-experimental";

    public static bool IsExperimentalBridgeEnabled()
        => string.Equals(
            Environment.GetEnvironmentVariable(EnvVarName),
            ExperimentalValue,
            StringComparison.OrdinalIgnoreCase);

    internal static bool CanUseInstalledRuntime(string? format = null, string? operation = null)
    {
        var runtime = HwpRuntimeProbe.Probe();
        return IsBridgeFormat(format) && OperationRuntimeAvailable(runtime, operation);
    }

    /// <summary>
    /// Returns the active engine. Throws HwpEngineException with bridge_missing
    /// if rhwp-experimental is requested but the bridge executable is not found.
    /// </summary>
    public static IHwpEngine GetEngine(string? format = null, string? operation = null)
    {
        var runtime = HwpRuntimeProbe.Probe();
        var explicitBridge = IsExperimentalBridgeEnabled();
        var autoRuntime = IsBridgeFormat(format) && OperationRuntimeAvailable(runtime, operation);
        if (!explicitBridge && !autoRuntime)
        {
            if (string.Equals(format, HwpCapabilityConstants.FormatHwp, StringComparison.OrdinalIgnoreCase)
                && IsBridgeBackedOperation(operation))
            {
                throw new HwpEngineException(
                    "Binary .hwp operation requires packaged rhwp sidecars or OFFICECLI_HWP_ENGINE=rhwp-experimental.",
                    HwpCapabilityConstants.ReasonBridgeNotEnabled,
                    "Run ./dev-install.sh or set OFFICECLI_RHWP_BRIDGE_PATH and OFFICECLI_RHWP_API_BIN.",
                    [],
                    format,
                    operation,
                    engine: HwpCapabilityConstants.EngineNone,
                    engineMode: HwpCapabilityConstants.ModeNone);
            }
            return new CustomHwpxEngine();
        }

        if (runtime.BridgePath == null)
            throw new HwpEngineException(
                "rhwp-officecli-bridge is not available.",
                HwpCapabilityConstants.ReasonBridgeMissing,
                "Run `officecli hwp doctor --json`; install sidecars with ./dev-install.sh or set OFFICECLI_RHWP_BRIDGE_PATH.",
                [],
                format,
                operation,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);

        var runtimeReason = explicitBridge ? null : RuntimeBlockedReason(runtime, operation);
        if (runtimeReason != null)
            throw new HwpEngineException(
                $"rhwp runtime is not available for operation '{operation ?? "unknown"}'.",
                runtimeReason,
                "Run `officecli hwp doctor --json`; install sidecars with ./dev-install.sh or set explicit rhwp environment paths.",
                [],
                format,
                operation,
                engine: HwpCapabilityConstants.EngineRhwpBridge,
                engineMode: HwpCapabilityConstants.ModeExperimental);

        var bridge = RhwpBridgeEngine.TryCreate(out var missingReason);
        if (bridge != null)
            return bridge;

        throw new HwpEngineException(
            $"OFFICECLI_HWP_ENGINE=rhwp-experimental requested but bridge not available: {missingReason}",
            HwpCapabilityConstants.ReasonBridgeMissing,
            "Run `officecli help hwp`; set OFFICECLI_RHWP_BRIDGE_PATH and, for fields/text/table mutation, OFFICECLI_RHWP_API_BIN.",
            [],
            format,
            operation,
            engine: HwpCapabilityConstants.EngineRhwpBridge,
            engineMode: HwpCapabilityConstants.ModeExperimental);
    }

    private static bool IsBridgeFormat(string? format)
        => string.Equals(format, HwpCapabilityConstants.FormatHwp, StringComparison.OrdinalIgnoreCase)
            || string.Equals(format, HwpCapabilityConstants.FormatHwpx, StringComparison.OrdinalIgnoreCase);

    private static bool IsBridgeBackedOperation(string? operation)
        => operation is HwpCapabilityConstants.OperationReadText
            or HwpCapabilityConstants.OperationRenderSvg
            or HwpCapabilityConstants.OperationListFields
            or HwpCapabilityConstants.OperationReadField
            or HwpCapabilityConstants.OperationFillField
            or HwpCapabilityConstants.OperationReplaceText
            or HwpCapabilityConstants.OperationSetTableCell
            or HwpCapabilityConstants.OperationSaveAsHwp;

    private static bool OperationRuntimeAvailable(HwpRuntimeProbeResult runtime, string? operation)
        => operation switch
        {
            HwpCapabilityConstants.OperationReadText or HwpCapabilityConstants.OperationRenderSvg
                => runtime.ReadRenderAvailable,
            HwpCapabilityConstants.OperationListFields
                or HwpCapabilityConstants.OperationReadField
                or HwpCapabilityConstants.OperationFillField
                or HwpCapabilityConstants.OperationReplaceText
                or HwpCapabilityConstants.OperationSetTableCell
                or HwpCapabilityConstants.OperationSaveAsHwp
                => runtime.MutationAvailable,
            _ => runtime.BridgeAvailable
        };

    private static string? RuntimeBlockedReason(HwpRuntimeProbeResult runtime, string? operation)
    {
        if (OperationRuntimeAvailable(runtime, operation))
            return null;
        if (!runtime.BridgeAvailable)
            return HwpCapabilityConstants.ReasonBridgeMissing;
        if (operation is HwpCapabilityConstants.OperationReadText or HwpCapabilityConstants.OperationRenderSvg
            && !runtime.ApiAvailable && !runtime.RhwpAvailable)
            return HwpCapabilityConstants.ReasonRhwpRuntimeMissing;
        if (!runtime.ApiAvailable)
            return HwpCapabilityConstants.ReasonRhwpApiMissing;
        return HwpCapabilityConstants.ReasonBridgeMissing;
    }
}
