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
                ["ready"] = IsReady(capability),
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
                    or HwpCapabilityConstants.ReasonBridgeMissing))
        {
            yield return "export OFFICECLI_HWP_ENGINE=rhwp-experimental";
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
            case HwpCapabilityConstants.OperationSetTableCell:
                yield return "section";
                yield return "parent-para";
                yield return "control";
                yield return "cell";
                yield return "value";
                yield return "output";
                break;
        }
    }

    private static string? BuildExample(string operation)
        => operation switch
        {
            HwpCapabilityConstants.OperationReadText => "officecli view input.hwp text --json",
            HwpCapabilityConstants.OperationRenderSvg => "officecli view input.hwp svg --page 1 --json",
            HwpCapabilityConstants.OperationListFields => "officecli view input.hwp fields --json",
            HwpCapabilityConstants.OperationReadField => "officecli view input.hwp field --field-name 회사명 --json",
            HwpCapabilityConstants.OperationFillField => "officecli set input.hwp /field --prop name=회사명 --prop value=리지 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationReplaceText => "officecli set input.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=output.hwp --json",
            HwpCapabilityConstants.OperationSetTableCell => "officecli set input.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=output.hwp --json",
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
