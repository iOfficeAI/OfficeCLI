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

    /// <summary>
    /// Returns the active engine. Throws HwpEngineException with bridge_missing
    /// if rhwp-experimental is requested but the bridge executable is not found.
    /// </summary>
    public static IHwpEngine GetEngine(string? format = null, string? operation = null)
    {
        if (!IsExperimentalBridgeEnabled())
            return new CustomHwpxEngine();

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
}
