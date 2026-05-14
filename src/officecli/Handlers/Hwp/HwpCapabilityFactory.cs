// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Handlers.Hwp;

public static class HwpCapabilityFactory
{
    public static HwpCapabilityReport BuildReport(string? customEngineVersion = null)
    {
        var version = customEngineVersion ?? $"officecli:{Assembly.GetExecutingAssembly().GetName().Version}";
        var runtime = HwpRuntimeProbe.Probe();
        var formats = new Dictionary<string, HwpFormatCapability>
        {
            [HwpCapabilityConstants.FormatHwpx] = BuildHwpx(version, runtime),
            [HwpCapabilityConstants.FormatHwp] = BuildHwp(runtime)
        };

        return new HwpCapabilityReport(
            HwpCapabilityConstants.SchemaVersion,
            Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown",
            DateTimeOffset.UtcNow,
            formats);
    }

    private static HwpFormatCapability BuildHwpx(string engineVersion, HwpRuntimeProbeResult runtime)
    {
        var apiBlockedReason = ApiBlockedReason(runtime);
        var readRenderBlockedReason = ReadRenderBlockedReason(runtime);
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = ExperimentalCustom(engineVersion,
                ["docs/hwpx-current-operation-inventory.md"]),
            [HwpCapabilityConstants.OperationRenderSvg] = runtime.ReadRenderAvailable
                ? ExperimentalBridge([ "tests/golden/hwp/rhwp-smoke/officecli-view/hwpx-svg.pretty.json" ])
                : ExperimentalBridgeBlocked(readRenderBlockedReason),
            [HwpCapabilityConstants.OperationListFields] = runtime.MutationAvailable
                ? ExperimentalBridge([])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationReadField] = runtime.MutationAvailable
                ? ExperimentalBridge([])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationFillField] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationReplaceText] = runtime.MutationAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/replace-hwpx-government.json"])
                : ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationSetTableCell] = Unsupported(
                HwpCapabilityConstants.ReasonRoundTripUnverified),
            [HwpCapabilityConstants.OperationCreateBlank] = ExperimentalCustom(engineVersion,
                ["src/officecli/Resources/base.hwpx"]),
            [HwpCapabilityConstants.OperationSaveOriginal] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationSaveAsHwp] = runtime.MutationAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/main.rs"])
                : ExperimentalBridgeBlocked(apiBlockedReason)
        };

        return new HwpFormatCapability(
            HwpCapabilityConstants.StatusExperimental,
            HwpCapabilityConstants.WriteStatusOperationGated,
            HwpCapabilityConstants.EngineCustom,
            operations,
            runtime.ReadRenderAvailable || runtime.MutationAvailable
                ? ["HWPX default engine remains custom; rhwp bridge is used only for opt-in text/svg/field/text-replace paths."]
                : ["HWPX operations are advertised only after per-operation round-trip evidence exists."]);
    }

    private static HwpFormatCapability BuildHwp(HwpRuntimeProbeResult runtime)
    {
        var readRenderBlockedReason = ReadRenderBlockedReason(runtime);
        var apiBlockedReason = ApiBlockedReason(runtime);
        var createBlockedReason = runtime.ApiAvailable
            ? HwpCapabilityConstants.ReasonRoundTripUnverified
            : HwpCapabilityConstants.ReasonRhwpApiMissing;
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = runtime.ReadRenderAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-smoke/officecli-view/text.pretty.json"])
                : ExperimentalBridgeBlocked(readRenderBlockedReason),
            [HwpCapabilityConstants.OperationRenderSvg] = runtime.ReadRenderAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-smoke/officecli-view/svg.pretty.json"])
                : ExperimentalBridgeBlocked(readRenderBlockedReason),
            [HwpCapabilityConstants.OperationListFields] = runtime.MutationAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/field-list.json"])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationReadField] = runtime.MutationAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/field-read-company-name.json"])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationFillField] = runtime.MutationAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/field-set-company-name-cli-readback.json"])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationReplaceText] = runtime.MutationAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/replace-marketing-title.json"])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationSetTableCell] = runtime.MutationAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-tables/table-cell-set-news-title.json"])
                : ExperimentalBridgeBlocked(apiBlockedReason),
            [HwpCapabilityConstants.OperationCreateBlank] = runtime.CreateBlankAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/main.rs"])
                : ExperimentalBridgeBlocked(createBlockedReason),
            [HwpCapabilityConstants.OperationSaveOriginal] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden),
            [HwpCapabilityConstants.OperationSaveAsHwp] = runtime.MutationAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/main.rs"])
                : ExperimentalBridgeBlocked(apiBlockedReason)
        };

        return new HwpFormatCapability(
            runtime.ReadRenderAvailable
                ? HwpCapabilityConstants.StatusExperimental
                : HwpCapabilityConstants.StatusUnsupported,
            runtime.MutationAvailable || runtime.CreateBlankAvailable
                ? HwpCapabilityConstants.WriteStatusOperationGated
                : HwpCapabilityConstants.WriteStatusUnsupported,
            runtime.BridgeAvailable || runtime.ApiAvailable
                ? HwpCapabilityConstants.EngineRhwpBridge
                : HwpCapabilityConstants.EngineNone,
            operations,
            HwpWarnings(runtime));
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

    private static string ReadRenderBlockedReason(HwpRuntimeProbeResult runtime)
    {
        if (!runtime.EngineRequested && !runtime.BridgeAvailable && !runtime.ApiAvailable && !runtime.RhwpAvailable)
            return HwpCapabilityConstants.ReasonBridgeNotEnabled;
        if (!runtime.BridgeAvailable)
            return HwpCapabilityConstants.ReasonBridgeMissing;
        if (!runtime.ApiAvailable && !runtime.RhwpAvailable)
            return HwpCapabilityConstants.ReasonRhwpRuntimeMissing;
        return HwpCapabilityConstants.ReasonBridgeMissing;
    }

    private static string ApiBlockedReason(HwpRuntimeProbeResult runtime)
    {
        if (!runtime.EngineRequested && !runtime.BridgeAvailable && !runtime.ApiAvailable)
            return HwpCapabilityConstants.ReasonBridgeNotEnabled;
        if (!runtime.BridgeAvailable)
            return HwpCapabilityConstants.ReasonBridgeMissing;
        if (!runtime.ApiAvailable)
            return HwpCapabilityConstants.ReasonRhwpApiMissing;
        return HwpCapabilityConstants.ReasonBridgeMissing;
    }

    private static string[] HwpWarnings(HwpRuntimeProbeResult runtime)
    {
        if (runtime.MutationAvailable)
        {
            return [
                "Experimental rhwp bridge is available from installed sidecars or explicit environment paths.",
                "Default mutations require explicit --prop output=<path>.",
                "Do not claim production-grade HWP fidelity without fixture and Hancom round-trip evidence."
            ];
        }

        if (runtime.CreateBlankAvailable)
        {
            return [
                "Blank .hwp creation is available through rhwp-field-bridge.",
                "Read/render/mutation still require rhwp-officecli-bridge plus rhwp-field-bridge or rhwp CLI."
            ];
        }

        return [
            "Binary .hwp support requires packaged rhwp sidecars beside officecli or explicit environment paths.",
            SetupHint()
        ];
    }

    private static string SetupHint()
        => "Run ./dev-install.sh to install rhwp sidecars beside officecli, or set OFFICECLI_RHWP_BRIDGE_PATH/OFFICECLI_RHWP_API_BIN/OFFICECLI_RHWP_BIN.";
}
