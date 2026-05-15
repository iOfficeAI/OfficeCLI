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
        var readRenderBlockedReason = ReadRenderBlockedReason(runtime);
        var operations = new Dictionary<string, HwpOperationCapability>
        {
            [HwpCapabilityConstants.OperationReadText] = runtime.ReadRenderAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-smoke/hwpx-officecli-view/text.pretty.json"])
                : ExperimentalCustom(engineVersion, ["tests/golden/hwp/rhwp-smoke/hwpx-officecli-view/text.pretty.json"]),
            [HwpCapabilityConstants.OperationRenderSvg] = runtime.ReadRenderAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-smoke/hwpx-officecli-view/svg.pretty.json"])
                : ExperimentalBridgeBlocked(readRenderBlockedReason),
            [HwpCapabilityConstants.OperationRenderPng] = runtime.RenderPngAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "render-png")),
            [HwpCapabilityConstants.OperationExportPdf] = runtime.ExportPdfAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "export-pdf")),
            [HwpCapabilityConstants.OperationExportMarkdown] = runtime.ExportMarkdownAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "export-markdown")),
            [HwpCapabilityConstants.OperationThumbnail] = runtime.ThumbnailAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "thumbnail")),
            [HwpCapabilityConstants.OperationDocumentInfo] = runtime.DocumentInfoAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "document-info")),
            [HwpCapabilityConstants.OperationDiagnostics] = runtime.DiagnosticsAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "diagnostics")),
            [HwpCapabilityConstants.OperationDumpControls] = runtime.DumpControlsAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "dump-controls")),
            [HwpCapabilityConstants.OperationDumpPages] = runtime.DumpPagesAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "dump-pages")),
            [HwpCapabilityConstants.OperationListFields] = runtime.ListFieldsAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "list-fields")),
            [HwpCapabilityConstants.OperationReadField] = runtime.ReadFieldAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "get-field")),
            [HwpCapabilityConstants.OperationFillField] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationReplaceText] = runtime.ReplaceTextAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/officecli-replace-hwpx-government.json"])
                : ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationInsertText] = runtime.InsertTextAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_text.rs"])
                : ExperimentalBridgeBlocked(InsertTextBlockedReason(runtime)),
            [HwpCapabilityConstants.OperationSetTableCell] = runtime.SetTableCellAvailable
                ? ExperimentalBridge(["tests/OfficeCli.Tests/Hwp/HwpBridgeTableScanTests.cs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "set-cell-text")),
            [HwpCapabilityConstants.OperationCreateBlank] = ExperimentalCustom(engineVersion,
                ["src/officecli/Resources/base.hwpx"]),
            [HwpCapabilityConstants.OperationSaveOriginal] = ExperimentalCustom(engineVersion, []),
            [HwpCapabilityConstants.OperationNativeRead] = runtime.NativeOpAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_native.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "native-op")),
            [HwpCapabilityConstants.OperationNativeMutation] = runtime.NativeOpAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_native.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "native-op")),
            [HwpCapabilityConstants.OperationSaveAsHwp] = runtime.SaveAsHwpAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/main.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "save-as-hwp"))
        };

        return new HwpFormatCapability(
            HwpCapabilityConstants.StatusExperimental,
            HwpCapabilityConstants.WriteStatusOperationGated,
            HwpCapabilityConstants.EngineCustom,
            operations,
            runtime.ReadRenderAvailable || runtime.MutationAvailable
                ? ["HWPX default engine remains custom, and rhwp sidecars are used for operations that the bridge exposes."]
                : ["HWPX operations are advertised only after per-operation round-trip evidence exists."]);
    }

    private static HwpFormatCapability BuildHwp(HwpRuntimeProbeResult runtime)
    {
        var readRenderBlockedReason = ReadRenderBlockedReason(runtime);
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
            [HwpCapabilityConstants.OperationRenderPng] = runtime.RenderPngAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "render-png")),
            [HwpCapabilityConstants.OperationExportPdf] = runtime.ExportPdfAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "export-pdf")),
            [HwpCapabilityConstants.OperationExportMarkdown] = runtime.ExportMarkdownAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "export-markdown")),
            [HwpCapabilityConstants.OperationThumbnail] = runtime.ThumbnailAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "thumbnail")),
            [HwpCapabilityConstants.OperationDocumentInfo] = runtime.DocumentInfoAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "document-info")),
            [HwpCapabilityConstants.OperationDiagnostics] = runtime.DiagnosticsAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "diagnostics")),
            [HwpCapabilityConstants.OperationDumpControls] = runtime.DumpControlsAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "dump-controls")),
            [HwpCapabilityConstants.OperationDumpPages] = runtime.DumpPagesAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_view.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "dump-pages")),
            [HwpCapabilityConstants.OperationListFields] = runtime.ListFieldsAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/field-list.json"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "list-fields")),
            [HwpCapabilityConstants.OperationReadField] = runtime.ReadFieldAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/field-read-company-name.json"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "get-field")),
            [HwpCapabilityConstants.OperationFillField] = runtime.FillFieldAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/field-set-company-name-cli-readback.json"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "set-field")),
            [HwpCapabilityConstants.OperationReplaceText] = runtime.ReplaceTextAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-fields/officecli-replace-marketing-title.json"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "replace-text")),
            [HwpCapabilityConstants.OperationInsertText] = runtime.InsertTextAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_text.rs"])
                : ExperimentalBridgeBlocked(InsertTextBlockedReason(runtime)),
            [HwpCapabilityConstants.OperationReadTableCell] = runtime.ReadTableCellAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "get-cell-text")),
            [HwpCapabilityConstants.OperationScanCells] = runtime.ScanCellsAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "scan-cells")),
            [HwpCapabilityConstants.OperationSetTableCell] = runtime.SetTableCellAvailable
                ? ExperimentalBridge(["tests/golden/hwp/rhwp-tables/officecli-set-cell-hwp-table-readback.json"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "set-cell-text")),
            [HwpCapabilityConstants.OperationCreateBlank] = runtime.CreateBlankAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/main.rs"])
                : ExperimentalBridgeBlocked(createBlockedReason),
            [HwpCapabilityConstants.OperationSaveOriginal] = Unsupported(
                HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden),
            [HwpCapabilityConstants.OperationConvertToEditable] = runtime.ConvertToEditableAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "convert-to-editable")),
            [HwpCapabilityConstants.OperationNativeRead] = runtime.NativeOpAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_native.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "native-op")),
            [HwpCapabilityConstants.OperationNativeMutation] = runtime.NativeOpAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/ops_native.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "native-op")),
            [HwpCapabilityConstants.OperationSaveAsHwp] = runtime.SaveAsHwpAvailable
                ? ExperimentalBridge(["src/rhwp-field-bridge/src/main.rs"])
                : ExperimentalBridgeBlocked(ApiCommandBlockedReason(runtime, "save-as-hwp"))
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

    private static string InsertTextBlockedReason(HwpRuntimeProbeResult runtime)
    {
        if (!runtime.EngineRequested && !runtime.BridgeAvailable && !runtime.ApiAvailable)
            return HwpCapabilityConstants.ReasonBridgeNotEnabled;
        if (!runtime.BridgeAvailable)
            return HwpCapabilityConstants.ReasonBridgeMissing;
        if (!runtime.ApiAvailable || !runtime.InsertTextAvailable)
            return HwpCapabilityConstants.ReasonRhwpApiMissingOrTooOld;
        return HwpCapabilityConstants.ReasonBridgeMissing;
    }

    private static string ApiCommandBlockedReason(HwpRuntimeProbeResult runtime, string command)
    {
        if (!runtime.EngineRequested && !runtime.BridgeAvailable && !runtime.ApiAvailable)
            return HwpCapabilityConstants.ReasonBridgeNotEnabled;
        if (!runtime.BridgeAvailable)
            return HwpCapabilityConstants.ReasonBridgeMissing;
        if (!runtime.ApiAvailable || !runtime.ApiCommands.Contains(command))
            return HwpCapabilityConstants.ReasonRhwpApiMissingOrTooOld;
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
