// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Handlers.Hwp;

public static class HwpCapabilityFactory
{
    public static HwpCapabilityReport BuildReport(string? customEngineVersion = null)
    {
        var version = customEngineVersion ?? $"officecli:{Assembly.GetExecutingAssembly().GetName().Version}";
        var bridgeEnabled = HwpEngineSelector.IsExperimentalBridgeEnabled();
        string? missingReason = null;
        var bridgeAvailable = bridgeEnabled && RhwpBridgeEngine.TryCreate(out missingReason) != null;
        var formats = new Dictionary<string, HwpFormatCapability>
        {
            [HwpCapabilityConstants.FormatHwpx] = BuildHwpx(version, bridgeAvailable, missingReason),
            [HwpCapabilityConstants.FormatHwp] = BuildHwp(bridgeEnabled, bridgeAvailable, missingReason)
        };

        return new HwpCapabilityReport(
            HwpCapabilityConstants.SchemaVersion,
            Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown",
            DateTimeOffset.UtcNow,
            formats);
    }

    private static HwpFormatCapability BuildHwpx(
        string engineVersion,
        bool bridgeAvailable,
        string? missingReason)
    {
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = ExperimentalCustom(engineVersion,
                ["docs/hwpx-current-operation-inventory.md"]),
            [HwpCapabilityConstants.OperationRenderSvg] = bridgeAvailable
                ? ExperimentalBridge([ "tests/golden/hwp/rhwp-smoke/officecli-view/hwpx-svg.pretty.json" ])
                : ExperimentalBridgeBlocked(BridgeUnsupportedReason(missingReason)),
            [HwpCapabilityConstants.OperationListFields] = bridgeAvailable
                ? ExperimentalBridge([])
                : ExperimentalBridgeBlocked(BridgeUnsupportedReason(missingReason)),
            [HwpCapabilityConstants.OperationReadField] = bridgeAvailable
                ? ExperimentalBridge([])
                : ExperimentalBridgeBlocked(BridgeUnsupportedReason(missingReason)),
            [HwpCapabilityConstants.OperationFillField] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationReplaceText] = bridgeAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/replace-hwpx-government.json"])
                : ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationSetTableCell] = Unsupported(
                HwpCapabilityConstants.ReasonRoundTripUnverified),
            [HwpCapabilityConstants.OperationSaveOriginal] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationSaveAsHwp] = Unsupported(
                HwpCapabilityConstants.ReasonUnsupportedOperation)
        };

        return new HwpFormatCapability(
            HwpCapabilityConstants.StatusExperimental,
            HwpCapabilityConstants.WriteStatusOperationGated,
            HwpCapabilityConstants.EngineCustom,
            operations,
            bridgeAvailable
                ? ["HWPX default engine remains custom; rhwp bridge is used only for opt-in text/svg/field/text-replace paths."]
                : ["HWPX operations are advertised only after per-operation round-trip evidence exists."]);
    }

    private static HwpFormatCapability BuildHwp(bool bridgeEnabled, bool bridgeAvailable, string? missingReason)
    {
        if (bridgeAvailable)
            return BuildBridgeHwp();

        var blockedReason = bridgeEnabled
            ? HwpCapabilityConstants.ReasonBridgeMissing
            : HwpCapabilityConstants.ReasonBridgeNotEnabled;
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = ExperimentalBridgeBlocked(blockedReason),
            [HwpCapabilityConstants.OperationRenderSvg] = ExperimentalBridgeBlocked(blockedReason),
            [HwpCapabilityConstants.OperationListFields] = ExperimentalBridgeBlocked(blockedReason),
            [HwpCapabilityConstants.OperationReadField] = ExperimentalBridgeBlocked(blockedReason),
            [HwpCapabilityConstants.OperationFillField] = ExperimentalBridgeBlocked(blockedReason),
            [HwpCapabilityConstants.OperationReplaceText] = ExperimentalBridgeBlocked(blockedReason),
            [HwpCapabilityConstants.OperationSetTableCell] = ExperimentalBridgeBlocked(blockedReason),
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
            bridgeEnabled
                ? [$"rhwp bridge requested but unavailable: {missingReason}", SetupHint()]
                : ["Binary .hwp support requires OFFICECLI_HWP_ENGINE=rhwp-experimental.", SetupHint()]);
    }

    private static HwpFormatCapability BuildBridgeHwp()
    {
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-smoke/officecli-view/text.pretty.json"]),
            [HwpCapabilityConstants.OperationRenderSvg] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-smoke/officecli-view/svg.pretty.json"]),
            [HwpCapabilityConstants.OperationListFields] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-fields/field-list.json"]),
            [HwpCapabilityConstants.OperationReadField] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-fields/field-read-company-name.json"]),
            [HwpCapabilityConstants.OperationFillField] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-fields/field-set-company-name-cli-readback.json"]),
            [HwpCapabilityConstants.OperationReplaceText] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-fields/replace-marketing-title.json"]),
            [HwpCapabilityConstants.OperationSetTableCell] = ExperimentalBridge(
                ["tests/golden/hwp/rhwp-tables/table-cell-set-news-title.json"]),
            [HwpCapabilityConstants.OperationSaveOriginal] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden),
            [HwpCapabilityConstants.OperationSaveAsHwp] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden)
        };

        return new HwpFormatCapability(
            HwpCapabilityConstants.StatusExperimental,
            HwpCapabilityConstants.WriteStatusOperationGated,
            HwpCapabilityConstants.EngineRhwpBridge,
            operations,
            [
                "Experimental rhwp bridge is enabled. Mutations require explicit --prop output=<path>.",
                "Do not claim production-grade HWP fidelity without fixture and Hancom round-trip evidence."
            ]);
    }

    private static HwpOperationCapability ExperimentalCustom(string engineVersion, string[] evidence)
        => new(
            HwpOperationStatus.Experimental,
            HwpCapabilityConstants.EngineCustom,
            engineVersion,
            evidence,
            ["Not advertised until fixture round-trip and Hancom evidence are complete."],
            HwpCapabilityConstants.ReasonRoundTripUnverified);

    private static HwpOperationCapability ExperimentalBridge(string[] evidence)
        => new(
            HwpOperationStatus.Experimental,
            HwpCapabilityConstants.EngineRhwpBridge,
            "rhwp-experimental",
            evidence,
            ["Experimental rhwp bridge path; verify generated output before production use."],
            HwpCapabilityConstants.ReasonRoundTripUnverified);

    private static HwpOperationCapability ExperimentalBridgeBlocked(string reason)
        => new(
            HwpOperationStatus.Experimental,
            HwpCapabilityConstants.EngineRhwpBridge,
            null,
            [],
            [SetupHint()],
            reason);

    private static HwpOperationCapability Unsupported(string unsupportedReason)
        => new(
            HwpOperationStatus.Unsupported,
            HwpCapabilityConstants.EngineNone,
            null,
            [],
            [],
            unsupportedReason);

    private static string BridgeUnsupportedReason(string? missingReason)
        => missingReason == null
            ? HwpCapabilityConstants.ReasonBridgeNotEnabled
            : HwpCapabilityConstants.ReasonBridgeMissing;

    private static string SetupHint()
        => "Set OFFICECLI_HWP_ENGINE=rhwp-experimental, OFFICECLI_RHWP_BIN, OFFICECLI_RHWP_BRIDGE_PATH, and OFFICECLI_RHWP_API_BIN.";
}
