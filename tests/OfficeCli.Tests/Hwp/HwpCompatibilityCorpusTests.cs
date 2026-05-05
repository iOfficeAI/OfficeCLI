using System.Security.Cryptography;
using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public sealed class HwpCompatibilityCorpusTests
{
    private const string CompatibilityCorpusSchemaPath = "schemas/interfaces/compatibility-corpus.v1.schema.json";
    private const string ExpectedCapabilitiesSchemaPath = "schemas/interfaces/expected-capabilities.v1.schema.json";
    private const string ExpectedCapabilitiesPath = "tests/fixtures/common/expected-capabilities.json";

    private static readonly string[] ManifestPaths =
    [
        "tests/fixtures/hwp/manifest.json",
        "tests/fixtures/hwpx/manifest.json"
    ];

    private static readonly HashSet<string> KnownFormats = new(StringComparer.Ordinal)
    {
        HwpCapabilityConstants.FormatHwp,
        HwpCapabilityConstants.FormatHwpx
    };

    private static readonly HashSet<string> KnownCapabilityOperations = new(StringComparer.Ordinal)
    {
        HwpCapabilityConstants.OperationReadText,
        HwpCapabilityConstants.OperationRenderSvg,
        HwpCapabilityConstants.OperationListFields,
        HwpCapabilityConstants.OperationReadField,
        HwpCapabilityConstants.OperationFillField,
        HwpCapabilityConstants.OperationReplaceText,
        HwpCapabilityConstants.OperationSetTableCell,
        HwpCapabilityConstants.OperationSaveOriginal,
        HwpCapabilityConstants.OperationSaveAsHwp
    };

    private static readonly HashSet<string> KnownCorpusOperations = new(KnownCapabilityOperations, StringComparer.Ordinal)
    {
        "table_map"
    };

    private static readonly HashSet<string> KnownStatuses = new(StringComparer.Ordinal)
    {
        HwpCapabilityConstants.StatusUnsupported,
        HwpCapabilityConstants.StatusExperimental,
        HwpCapabilityConstants.StatusRoundTripVerified
    };

    private static readonly HashSet<string> KnownReasons = new(StringComparer.Ordinal)
    {
        HwpCapabilityConstants.ReasonUnsupportedFormat,
        HwpCapabilityConstants.ReasonUnsupportedOperation,
        HwpCapabilityConstants.ReasonUnsupportedEngine,
        HwpCapabilityConstants.ReasonRoundTripUnverified,
        HwpCapabilityConstants.ReasonBridgeNotEnabled,
        HwpCapabilityConstants.ReasonBridgeMissing,
        HwpCapabilityConstants.ReasonBridgeTimeout,
        HwpCapabilityConstants.ReasonBridgeInvalidJson,
        HwpCapabilityConstants.ReasonBridgeExitNonZero,
        HwpCapabilityConstants.ReasonBinaryHwpMutationForbidden,
        HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden,
        HwpCapabilityConstants.ReasonFixtureValidationFailed,
        HwpCapabilityConstants.ReasonCapabilitySchemaInvalid
    };

    [Theory]
    [InlineData("tests/fixtures/hwp/manifest.json", "hwp")]
    [InlineData("tests/fixtures/hwpx/manifest.json", "hwpx")]
    public void FixtureManifest_ReferencesExistingFilesAndEvidence(string manifestPath, string format)
    {
        var manifest = ReadJson(manifestPath);

        Assert.Equal(1, manifest["schemaVersion"]!.GetValue<int>());
        Assert.Equal(format, manifest["format"]!.GetValue<string>());

        foreach (var fixture in manifest["fixtures"]!.AsArray())
        {
            Assert.NotNull(fixture);
            var path = fixture!["path"]!.GetValue<string>();
            var fullPath = LocateRepoFile(path);
            Assert.Equal(fixture["sha256"]!.GetValue<string>(), Sha256File(fullPath));
            Assert.Equal(fixture["sizeBytes"]!.GetValue<long>(), new FileInfo(fullPath).Length);
            Assert.NotEmpty(fixture["classes"]!.AsArray());
            Assert.NotEmpty(fixture["verifiedOperations"]!.AsArray());

            foreach (var evidence in fixture["evidence"]!.AsArray())
                _ = LocateRepoFile(evidence!.GetValue<string>());
        }
    }

    [Fact]
    public void ManifestConformsToSchema()
    {
        AssertSchemaFileHasRequiredFields(CompatibilityCorpusSchemaPath, "format", "fixtures");

        foreach (var path in ManifestPaths)
            ValidateManifestContract(ReadJson(path));
    }

    [Fact]
    public void ExpectedCapabilities_ReferenceKnownOperationsAndEvidence()
    {
        var manifest = ReadJson(ExpectedCapabilitiesPath);

        foreach (var format in manifest["formats"]!.AsObject())
        {
            Assert.True(
                KnownFormats.Contains(format.Key),
                $"Unknown corpus format '{format.Key}'.");
            var operations = format.Value!["operations"]!.AsObject();
            Assert.NotEmpty(operations);
            foreach (var operation in operations)
            {
                Assert.Contains(operation.Key, KnownCapabilityOperations);
                var status = operation.Value!["status"]!.GetValue<string>();
                Assert.True(
                    KnownStatuses.Contains(status),
                    $"Unknown status '{status}' for {format.Key}.{operation.Key}.");
                foreach (var evidence in operation.Value["evidence"]!.AsArray())
                    _ = LocateRepoFile(evidence!.GetValue<string>());
            }
        }
    }

    [Fact]
    public void ExpectedCapabilitiesConformsToSchema()
    {
        AssertSchemaFileHasRequiredFields(ExpectedCapabilitiesSchemaPath, "formats", "operations");

        ValidateExpectedCapabilitiesContract(ReadJson(ExpectedCapabilitiesPath));
    }

    [Fact]
    public void BlockedOperationsRequireReasonAndNoEvidenceUnlessExplained()
    {
        var expectedCapabilities = ReadJson(ExpectedCapabilitiesPath);
        var observedHwpxTableCellBlock = false;

        foreach (var path in ManifestPaths)
        {
            var manifest = ReadJson(path);
            var format = RequireString(manifest, "format");
            foreach (var fixture in GetArray(manifest, "fixtures"))
            {
                foreach (var blockedOperation in GetArray(fixture!, "blockedOperations", required: false))
                {
                    var operation = RequireString(blockedOperation!, "operation");
                    var reason = RequireString(blockedOperation!, "reason");
                    _ = RequireString(blockedOperation!, "notes");
                    Assert.Contains(operation, KnownCorpusOperations);
                    Assert.Contains(reason, KnownReasons);

                    var operationCapability = expectedCapabilities["formats"]![format]!["operations"]![operation]!.AsObject();
                    Assert.Equal(HwpCapabilityConstants.StatusUnsupported, RequireString(operationCapability, "status"));
                    Assert.Equal(reason, RequireString(operationCapability, "reason"));
                    Assert.Empty(GetArray(operationCapability, "evidence"));

                    observedHwpxTableCellBlock |= format == HwpCapabilityConstants.FormatHwpx
                        && operation == HwpCapabilityConstants.OperationSetTableCell
                        && reason == HwpCapabilityConstants.ReasonRoundTripUnverified;
                }
            }
        }

        Assert.True(observedHwpxTableCellBlock, "HWPX set_table_cell must remain blocked as roundtrip_unverified.");
    }

    [Fact]
    public void ManifestSchemaRejectsUnknownFormat()
    {
        var manifest = Clone(ReadJson("tests/fixtures/hwp/manifest.json"));
        manifest["format"] = "odt";

        var error = Assert.Throws<InvalidOperationException>(() => ValidateManifestContract(manifest));
        Assert.Contains("Unknown corpus format", error.Message);
    }

    [Fact]
    public void ExpectedCapabilitiesSchemaRejectsUnknownOperation()
    {
        var capabilities = Clone(ReadJson(ExpectedCapabilitiesPath));
        capabilities["formats"]!["hwp"]!["operations"]!["unknown_operation"] = new JsonObject
        {
            ["status"] = HwpCapabilityConstants.StatusExperimental,
            ["evidence"] = new JsonArray()
        };

        var error = Assert.Throws<InvalidOperationException>(() => ValidateExpectedCapabilitiesContract(capabilities));
        Assert.Contains("Unknown capability operation", error.Message);
    }

    private static void ValidateManifestContract(JsonNode manifest)
    {
        Assert.Equal(1, RequireInt(manifest, "schemaVersion"));
        var format = RequireString(manifest, "format");
        if (!KnownFormats.Contains(format))
            throw new InvalidOperationException($"Unknown corpus format '{format}'.");
        _ = RequireString(manifest, "description");

        foreach (var fixture in GetArray(manifest, "fixtures"))
        {
            _ = RequireString(fixture!, "id");
            _ = RequireString(fixture!, "path");
            var sha256 = RequireString(fixture!, "sha256");
            Assert.True(IsLowercaseSha256(sha256), $"Fixture sha256 is invalid: {sha256}");
            Assert.True(RequireLong(fixture!, "sizeBytes") > 0);
            Assert.NotEmpty(GetStringArray(fixture!, "classes"));
            Assert.NotEmpty(GetStringArray(fixture!, "evidence"));
            _ = RequireString(fixture!, "notes");

            foreach (var operation in GetStringArray(fixture!, "verifiedOperations"))
            {
                if (!KnownCorpusOperations.Contains(operation))
                    throw new InvalidOperationException($"Unknown corpus operation '{operation}'.");
            }

            foreach (var blockedOperation in GetArray(fixture!, "blockedOperations", required: false))
            {
                var operation = RequireString(blockedOperation!, "operation");
                var reason = RequireString(blockedOperation!, "reason");
                _ = RequireString(blockedOperation!, "notes");
                Assert.Contains(operation, KnownCorpusOperations);
                Assert.Contains(reason, KnownReasons);
            }
        }
    }

    private static void ValidateExpectedCapabilitiesContract(JsonNode manifest)
    {
        Assert.Equal(1, RequireInt(manifest, "schemaVersion"));
        _ = RequireString(manifest, "description");
        var formats = manifest["formats"]!.AsObject();
        Assert.Equal(KnownFormats.OrderBy(static value => value), formats.Select(static format => format.Key).OrderBy(static value => value));

        foreach (var format in formats)
        {
            _ = RequireString(format.Value!, "defaultEngine");
            var operations = format.Value!["operations"]!.AsObject();
            Assert.NotEmpty(operations);
            foreach (var operation in operations)
            {
                if (!KnownCapabilityOperations.Contains(operation.Key))
                    throw new InvalidOperationException($"Unknown capability operation '{operation.Key}'.");
                var status = RequireString(operation.Value!, "status");
                Assert.Contains(status, KnownStatuses);
                var evidence = GetArray(operation.Value!, "evidence");
                foreach (var item in evidence)
                    _ = RequireString(item!);
                if (status == HwpCapabilityConstants.StatusUnsupported)
                {
                    var reason = RequireString(operation.Value!, "reason");
                    Assert.Contains(reason, KnownReasons);
                }
            }
        }
    }

    private static void AssertSchemaFileHasRequiredFields(string relativePath, params string[] fieldNames)
    {
        var schema = ReadJson(relativePath);
        Assert.Equal("object", RequireString(schema, "type"));
        var schemaText = schema.ToJsonString();
        foreach (var fieldName in fieldNames)
            Assert.Contains(fieldName, schemaText, StringComparison.Ordinal);
    }

    private static JsonNode ReadJson(string relativePath)
        => JsonNode.Parse(File.ReadAllText(LocateRepoFile(relativePath)))!;

    private static JsonNode Clone(JsonNode node)
        => JsonNode.Parse(node.ToJsonString())!;

    private static JsonArray GetArray(JsonNode node, string propertyName, bool required = true)
    {
        var value = node[propertyName];
        if (value is null)
        {
            if (!required) return new JsonArray();
            throw new InvalidOperationException($"Missing required array '{propertyName}'.");
        }
        return value.AsArray();
    }

    private static string[] GetStringArray(JsonNode node, string propertyName)
        => GetArray(node, propertyName).Select(static item => RequireString(item!)).ToArray();

    private static int RequireInt(JsonNode node, string propertyName)
        => node[propertyName]?.GetValue<int>() ?? throw new InvalidOperationException($"Missing required integer '{propertyName}'.");

    private static long RequireLong(JsonNode node, string propertyName)
        => node[propertyName]?.GetValue<long>() ?? throw new InvalidOperationException($"Missing required integer '{propertyName}'.");

    private static string RequireString(JsonNode node, string propertyName)
        => RequireString(node[propertyName] ?? throw new InvalidOperationException($"Missing required string '{propertyName}'."));

    private static string RequireString(JsonNode node)
    {
        var value = node.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value))
            throw new InvalidOperationException("Expected a non-empty string.");
        return value;
    }

    private static bool IsLowercaseSha256(string value)
        => value.Length == 64 && value.All(static character => character is >= '0' and <= '9' or >= 'a' and <= 'f');

    private static string LocateRepoFile(string relativePath)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, relativePath);
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new FileNotFoundException($"Required repo file was not found: {relativePath}");
    }

    private static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
}
