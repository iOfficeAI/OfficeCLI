// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;

namespace OfficeCli.Handlers.Hwp;

public enum HwpFormat
{
    Hwp,
    Hwpx
}

public enum HwpEngineMode
{
    None,
    Default,
    Experimental
}

public enum HwpOperationStatus
{
    Unsupported,
    Experimental,
    RoundTripVerified
}

public static class HwpCapabilityConstants
{
    public const int SchemaVersion = 1;

    public const string FormatHwp = "hwp";
    public const string FormatHwpx = "hwpx";

    public const string EngineCustom = "custom";
    public const string EngineRhwpBridge = "rhwp-bridge";
    public const string EngineNone = "none";

    public const string ModeNone = "none";
    public const string ModeDefault = "default";
    public const string ModeExperimental = "experimental";

    public const string StatusUnsupported = "unsupported";
    public const string StatusExperimental = "experimental";
    public const string StatusRoundTripVerified = "roundtrip-verified";

    public const string WriteStatusUnsupported = "unsupported";
    public const string WriteStatusOperationGated = "operation-gated";

    public const string OperationReadText = "read_text";
    public const string OperationRenderSvg = "render_svg";
    public const string OperationListFields = "list_fields";
    public const string OperationReadField = "read_field";
    public const string OperationFillField = "fill_field";
    public const string OperationReplaceText = "replace_text";
    public const string OperationSaveOriginal = "save_original";
    public const string OperationSaveAsHwp = "save_as_hwp";

    public const string ReasonUnsupportedFormat = "unsupported_format";
    public const string ReasonUnsupportedOperation = "unsupported_operation";
    public const string ReasonUnsupportedEngine = "unsupported_engine";
    public const string ReasonRoundTripUnverified = "roundtrip_unverified";
    public const string ReasonBridgeNotEnabled = "bridge_not_enabled";
    public const string ReasonBridgeMissing = "bridge_missing";
    public const string ReasonBridgeTimeout = "bridge_timeout";
    public const string ReasonBridgeInvalidJson = "bridge_invalid_json";
    public const string ReasonBridgeExitNonZero = "bridge_exit_nonzero";
    public const string ReasonBinaryHwpMutationForbidden = "binary_hwp_mutation_forbidden";
    public const string ReasonBinaryHwpWriteForbidden = "binary_hwp_write_forbidden";
    public const string ReasonFixtureValidationFailed = "fixture_validation_failed";
    public const string ReasonCapabilitySchemaInvalid = "capability_schema_invalid";
}

public sealed record HwpOperationCapability(
    HwpOperationStatus Status,
    string Engine,
    string? EngineVersion,
    string[] Evidence,
    string[] Warnings,
    string? UnsupportedReason
);

public sealed record HwpFormatCapability(
    string ReadStatus,
    string WriteStatus,
    string DefaultEngine,
    IReadOnlyDictionary<string, HwpOperationCapability> Operations,
    string[] Warnings
);

public sealed record HwpCapabilityReport(
    int SchemaVersion,
    string OfficeCliVersion,
    DateTimeOffset GeneratedAt,
    IReadOnlyDictionary<string, HwpFormatCapability> Formats
);

public sealed record HwpReadRequest(HwpFormat Format, string InputPath, long InputSizeBytes, bool Json);
public sealed record HwpTextPage(int Page, string Text);
public sealed record HwpTextResult(
    string Text,
    IReadOnlyList<HwpTextPage> Pages,
    string Engine,
    string? EngineVersion,
    string[] Evidence,
    string[] Warnings
);

public sealed record HwpRenderRequest(
    HwpFormat Format,
    string InputPath,
    string OutputDirectory,
    string PageSelector,
    long InputSizeBytes,
    bool Json
);

public sealed record HwpRenderedPage(int Page, string SvgPath, string Sha256);
public sealed record HwpRenderResult(
    IReadOnlyList<HwpRenderedPage> Pages,
    string ManifestPath,
    string Engine,
    string? EngineVersion,
    string[] Evidence,
    string[] Warnings
);

public sealed record HwpFieldListRequest(HwpFormat Format, string InputPath, long InputSizeBytes, bool Json);
public sealed record HwpFieldReadRequest(
    HwpFormat Format,
    string InputPath,
    string? FieldName,
    int? FieldId,
    long InputSizeBytes,
    bool Json
);
public sealed record HwpFieldListResult(
    JsonObject Fields,
    string Engine,
    string? EngineVersion,
    string[] Evidence,
    string[] Warnings
);
public sealed record HwpFieldReadResult(
    JsonObject Field,
    string Engine,
    string? EngineVersion,
    string[] Evidence,
    string[] Warnings
);

public sealed record HwpFillFieldRequest(
    HwpFormat Format,
    string InputPath,
    string OutputPath,
    IReadOnlyDictionary<string, string> Fields,
    bool Json
)
{
    public IReadOnlyDictionary<int, string> FieldIds { get; init; } = new Dictionary<int, string>();
}

public sealed record HwpReplaceTextRequest(
    HwpFormat Format,
    string InputPath,
    string OutputPath,
    string Query,
    string Value,
    string Mode,
    bool CaseSensitive,
    bool Json
);

public sealed record HwpSaveOriginalRequest(HwpFormat Format, string InputPath, string OutputPath, bool Json);
public sealed record HwpSaveAsHwpRequest(HwpFormat Format, string InputPath, string OutputPath, bool Json);
public sealed record HwpMutationResult(
    string OutputPath,
    string Engine,
    string? EngineVersion,
    string[] Evidence,
    string[] Warnings
);
