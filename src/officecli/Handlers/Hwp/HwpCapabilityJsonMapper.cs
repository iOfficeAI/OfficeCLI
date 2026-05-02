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
            formats[format] = BuildFormatCapability(capability);

        return new JsonObject
        {
            ["schemaVersion"] = report.SchemaVersion,
            ["officecliVersion"] = report.OfficeCliVersion,
            ["generatedAt"] = report.GeneratedAt.ToString("O"),
            ["formats"] = formats
        };
    }

    private static JsonObject BuildFormatCapability(HwpFormatCapability capability)
    {
        var operations = new JsonObject();
        foreach (var (operation, opCapability) in capability.Operations)
            operations[operation] = BuildOperationCapability(opCapability);

        return new JsonObject
        {
            ["readStatus"] = capability.ReadStatus,
            ["writeStatus"] = capability.WriteStatus,
            ["defaultEngine"] = capability.DefaultEngine,
            ["operations"] = operations,
            ["warnings"] = ToJsonArray(capability.Warnings)
        };
    }

    private static JsonObject BuildOperationCapability(HwpOperationCapability capability)
    {
        return new JsonObject
        {
            ["status"] = StatusToJson(capability.Status),
            ["engine"] = capability.Engine,
            ["engineVersion"] = capability.EngineVersion,
            ["evidence"] = ToJsonArray(capability.Evidence),
            ["warnings"] = ToJsonArray(capability.Warnings),
            ["unsupportedReason"] = capability.UnsupportedReason
        };
    }

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
