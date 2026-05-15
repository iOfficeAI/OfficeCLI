// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;

namespace OfficeCli.Handlers.Hwp;

public static class HwpCapabilityJsonMapper
{
    public static JsonObject BuildEnvelope(HwpCapabilityReport report)
    {
        return new JsonObject
        {
            ["success"] = true,
            ["data"] = BuildData(report),
            ["warnings"] = new JsonArray()
        };
    }

    public static JsonObject BuildData(HwpCapabilityReport report)
    {
        var formats = new JsonObject();
        foreach (var (format, capability) in report.Formats)
            formats[format] = BuildFormatCapability(format, capability);

        return new JsonObject
        {
            ["schemaVersion"] = report.SchemaVersion,
            ["officecliVersion"] = report.OfficeCliVersion,
            ["generatedAt"] = report.GeneratedAt.ToString("O"),
            ["formats"] = formats
        };
    }

    private static JsonObject BuildFormatCapability(string format, HwpFormatCapability capability)
    {
        var operations = new JsonObject();
        foreach (var (operation, opCapability) in capability.Operations)
            operations[operation] = BuildOperationCapability(format, operation, opCapability);

        return new JsonObject
        {
            ["readStatus"] = capability.ReadStatus,
            ["writeStatus"] = capability.WriteStatus,
            ["defaultEngine"] = capability.DefaultEngine,
            ["setupHints"] = ToJsonArray(BuildSetupHints(capability)),
            ["operations"] = operations,
            ["warnings"] = ToJsonArray(capability.Warnings)
        };
    }

    private static JsonObject BuildOperationCapability(string format, string operation, HwpOperationCapability capability)
    {
        var result = new JsonObject
        {
            ["status"] = StatusToJson(capability.Status),
            ["support"] = StatusToJson(capability.Status),
            ["ready"] = IsReady(capability),
            ["engine"] = capability.Engine,
            ["engineVersion"] = capability.EngineVersion,
            ["evidence"] = ToJsonArray(capability.Evidence),
            ["warnings"] = ToJsonArray(capability.Warnings),
            ["unsupportedReason"] = capability.UnsupportedReason,
            ["blockedBy"] = ToJsonArray(BuildBlockedBy(capability)),
            ["requiredArgs"] = ToJsonArray(BuildRequiredArgs(operation)),
            ["example"] = BuildExample(operation)
        };
        if (format == HwpCapabilityConstants.FormatHwp
            && operation == HwpCapabilityConstants.OperationReplaceText)
        {
            result["safeInPlace"] = new JsonObject
            {
                ["support"] = HwpCapabilityConstants.StatusExperimental,
                ["ready"] = IsSafeInPlaceReady(capability),
                ["requires"] = ToJsonArray(["--in-place", "--backup", "--verify"]),
                ["example"] = "officecli set input.hwp /text --prop find=마케팅 --prop value=브릿지 --in-place --backup --verify --json",
                ["policy"] = "creates temp output, provider readback, semantic delta, backup, manifest, then atomic replace"
            };
        }
        return result;
    }

    private static bool IsReady(HwpOperationCapability capability)
    {
        if (capability.Status == HwpOperationStatus.Unsupported)
            return false;
        return capability.UnsupportedReason is null
            or HwpCapabilityConstants.ReasonRoundTripUnverified;
    }

    private static bool IsSafeInPlaceReady(HwpOperationCapability capability)
    {
        if (!IsReady(capability))
            return false;

        return HwpRuntimeProbe.Probe().MutationAvailable;
    }

    private static IEnumerable<string> BuildBlockedBy(HwpOperationCapability capability)
    {
        if (capability.UnsupportedReason is null
            or HwpCapabilityConstants.ReasonRoundTripUnverified)
            yield break;
        yield return capability.UnsupportedReason;
    }

    private static IEnumerable<string> BuildSetupHints(HwpFormatCapability capability)
    {
        if (capability.Operations.Values.Any(op =>
                op.UnsupportedReason is HwpCapabilityConstants.ReasonBridgeNotEnabled
                    or HwpCapabilityConstants.ReasonBridgeMissing
                    or HwpCapabilityConstants.ReasonRhwpRuntimeMissing
                    or HwpCapabilityConstants.ReasonRhwpApiMissing
                    or HwpCapabilityConstants.ReasonRhwpApiMissingOrTooOld))
        {
            yield return "run ./dev-install.sh to install rhwp sidecars beside officecli";
            yield return "export OFFICECLI_RHWP_BIN=/path/to/rhwp";
            yield return "export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge.dll";
            yield return "export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge";
            yield return "officecli help hwp";
        }
    }

    private static IEnumerable<string> BuildRequiredArgs(string operation)
    {
        switch (operation)
        {
            case HwpCapabilityConstants.OperationReadField:
                yield return "field-name|field-id";
                break;
            case HwpCapabilityConstants.OperationFillField:
                yield return "name|id";
                yield return "value";
                yield return "output";
                break;
            case HwpCapabilityConstants.OperationReplaceText:
                yield return "find";
                yield return "value";
                yield return "output|--in-place";
                yield return "--backup when --in-place";
                yield return "--verify when --in-place";
                break;
            case HwpCapabilityConstants.OperationInsertText:
                yield return "value";
                yield return "output";
                yield return "section? default 0";
                yield return "paragraph? default 0";
                yield return "offset? default 0";
                break;
            case HwpCapabilityConstants.OperationRenderPng:
            case HwpCapabilityConstants.OperationExportMarkdown:
            case HwpCapabilityConstants.OperationDumpPages:
                yield return "page? default all";
                break;
            case HwpCapabilityConstants.OperationExportPdf:
                yield return "output";
                yield return "page? default all";
                break;
            case HwpCapabilityConstants.OperationDumpControls:
                yield break;
            case HwpCapabilityConstants.OperationThumbnail:
                yield return "output";
                break;
            case HwpCapabilityConstants.OperationReadTableCell:
                yield return "section";
                yield return "parent-para";
                yield return "control";
                yield return "cell";
                yield return "cell-para";
                break;
            case HwpCapabilityConstants.OperationScanCells:
                yield return "section? default 0";
                yield return "max-parent-para? default 50";
                break;
            case HwpCapabilityConstants.OperationSetTableCell:
                yield return "section";
                yield return "parent-para";
                yield return "control";
                yield return "cell";
                yield return "value";
                yield return "output";
                break;
            case HwpCapabilityConstants.OperationConvertToEditable:
            case HwpCapabilityConstants.OperationSaveAsHwp:
                yield return "output";
                break;
            case HwpCapabilityConstants.OperationNativeRead:
                yield return "op";
                yield return "native-arg? key=value";
                break;
            case HwpCapabilityConstants.OperationNativeMutation:
                yield return "op";
                yield return "output";
                yield return "operation-specific props";
                break;
        }
    }

    private static string? BuildExample(string operation)
        => operation switch
        {
            HwpCapabilityConstants.OperationReadText => "officecli view input.hwp text --json",
            HwpCapabilityConstants.OperationRenderSvg => "officecli view input.hwp svg --page 1 --json",
            HwpCapabilityConstants.OperationRenderPng => "officecli view input.hwp png --page 1 --out /tmp/hwp-png --json",
            HwpCapabilityConstants.OperationExportPdf => "officecli view input.hwp pdf --page 1 --out output.pdf --json",
            HwpCapabilityConstants.OperationExportMarkdown => "officecli view input.hwp markdown --json",
            HwpCapabilityConstants.OperationThumbnail => "officecli view input.hwp thumbnail --out thumb.png --json",
            HwpCapabilityConstants.OperationDocumentInfo => "officecli view input.hwp info --json",
            HwpCapabilityConstants.OperationDiagnostics => "officecli view input.hwp diagnostics --json",
            HwpCapabilityConstants.OperationDumpControls => "officecli view input.hwp dump --json",
            HwpCapabilityConstants.OperationDumpPages => "officecli view input.hwp pages --page 1 --json",
            HwpCapabilityConstants.OperationListFields => "officecli view input.hwp fields --json",
            HwpCapabilityConstants.OperationReadField => "officecli view input.hwp field --field-name 회사명 --json",
            HwpCapabilityConstants.OperationFillField => "officecli set input.hwp /field --prop name=회사명 --prop value=리지 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationReplaceText => "officecli set input.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationInsertText => "officecli add input.hwp /text --type text --prop value=본문 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationReadTableCell => "officecli view input.hwp table-cell --section 0 --parent-para 3 --control 0 --cell 0 --cell-para 0 --json",
            HwpCapabilityConstants.OperationScanCells => "officecli view input.hwp tables --section 0 --json",
            HwpCapabilityConstants.OperationSetTableCell => "officecli set input.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationCreateBlank => "officecli create output.hwp --json",
            HwpCapabilityConstants.OperationConvertToEditable => "officecli set input.hwp /convert-to-editable --prop output=editable.hwp --json",
            HwpCapabilityConstants.OperationNativeRead => "officecli view input.hwp native --op get-style-list --json",
            HwpCapabilityConstants.OperationNativeMutation => "officecli set input.hwp /native-op --prop op=split-paragraph --prop paragraph=0 --prop offset=5 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationSaveAsHwp => "officecli set input.hwpx /save-as-hwp --prop output=output.hwp --json",
            _ => null
        };

    private static string StatusToJson(HwpOperationStatus status)
    {
        return status switch
        {
            HwpOperationStatus.Unsupported => HwpCapabilityConstants.StatusUnsupported,
            HwpOperationStatus.Experimental => HwpCapabilityConstants.StatusExperimental,
            HwpOperationStatus.RoundTripVerified => HwpCapabilityConstants.StatusRoundTripVerified,
            _ => throw new ArgumentOutOfRangeException(nameof(status), status, "Unknown HWP operation status")
        };
    }

    internal static JsonArray ToJsonArray(IEnumerable<string> values)
    {
        var array = new JsonArray();
        foreach (var value in values)
            array.Add((JsonNode?)JsonValue.Create(value));
        return array;
    }
}
