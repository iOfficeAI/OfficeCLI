// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Handlers.Hwp;

public static class HwpCapabilityFactory
{
    public static HwpCapabilityReport BuildReport(string? customEngineVersion = null)
    {
        var version = customEngineVersion ?? $"officecli:{Assembly.GetExecutingAssembly().GetName().Version}";
        var formats = new Dictionary<string, HwpFormatCapability>
        {
            [HwpCapabilityConstants.FormatHwpx] = BuildHwpx(version),
            [HwpCapabilityConstants.FormatHwp] = BuildHwp()
        };

        return new HwpCapabilityReport(
            HwpCapabilityConstants.SchemaVersion,
            Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown",
            DateTimeOffset.UtcNow,
            formats);
    }

    private static HwpFormatCapability BuildHwpx(string engineVersion)
    {
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = ExperimentalCustom(engineVersion,
                ["docs/hwpx-current-operation-inventory.md"]),
            [HwpCapabilityConstants.OperationRenderSvg] = Unsupported(
                HwpCapabilityConstants.ReasonUnsupportedOperation),
            [HwpCapabilityConstants.OperationFillField] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationSaveOriginal] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationSaveAsHwp] = Unsupported(
                HwpCapabilityConstants.ReasonUnsupportedOperation)
        };

        return new HwpFormatCapability(
            HwpCapabilityConstants.StatusExperimental,
            HwpCapabilityConstants.WriteStatusOperationGated,
            HwpCapabilityConstants.EngineCustom,
            operations,
            ["HWPX operations are advertised only after per-operation round-trip evidence exists."]);
    }

    private static HwpFormatCapability BuildHwp()
    {
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = Unsupported(
                HwpCapabilityConstants.ReasonBridgeNotEnabled),
            [HwpCapabilityConstants.OperationRenderSvg] = Unsupported(
                HwpCapabilityConstants.ReasonBridgeNotEnabled),
            [HwpCapabilityConstants.OperationFillField] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpMutationForbidden),
            [HwpCapabilityConstants.OperationSaveOriginal] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden),
            [HwpCapabilityConstants.OperationSaveAsHwp] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden)
        };

        return new HwpFormatCapability(
            HwpCapabilityConstants.StatusUnsupported,
            HwpCapabilityConstants.WriteStatusUnsupported,
            HwpCapabilityConstants.EngineNone,
            operations,
            ["Binary .hwp edit/write/save-as is not advertised."]);
    }

    private static HwpOperationCapability ExperimentalCustom(string engineVersion, string[] evidence)
        => new(
            HwpOperationStatus.Experimental,
            HwpCapabilityConstants.EngineCustom,
            engineVersion,
            evidence,
            ["Not advertised until fixture round-trip and Hancom evidence are complete."],
            HwpCapabilityConstants.ReasonRoundTripUnverified);

    private static HwpOperationCapability Unsupported(string unsupportedReason)
        => new(
            HwpOperationStatus.Unsupported,
            HwpCapabilityConstants.EngineNone,
            null,
            [],
            [],
            unsupportedReason);
}
