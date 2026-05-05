using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using Json.Schema;
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

    private static readonly string[] RequiredFixtureClasses =
    [
        "multi-section",
        "merged-cell-tables",
        "nested-tables",
        "pictures-bindata",
        "headers-footers",
        "equations",
        "unicode-edge-cases",
        "malformed-hwpx-package"
    ];

    private static readonly HashSet<string> KnownCoverageStates = new(StringComparer.Ordinal)
    {
        "verified",
        "blocked",
        "external-manual"
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
        var schema = LoadSchema(CompatibilityCorpusSchemaPath);

        foreach (var path in ManifestPaths)
        {
            var manifest = ReadJson(path);
            AssertSchemaValid(schema, manifest);
            ValidateManifestContract(manifest);
        }
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
        var schema = LoadSchema(ExpectedCapabilitiesSchemaPath);
        var capabilities = ReadJson(ExpectedCapabilitiesPath);

        AssertSchemaValid(schema, capabilities);
        ValidateExpectedCapabilitiesContract(capabilities);
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
    public void RequiredFixtureClassesAreRepresented()
    {
        var observed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var path in ManifestPaths)
        {
            var manifest = ReadJson(path);
            var fixtureIds = new HashSet<string>(
                GetArray(manifest, "fixtures").Select(f => RequireString(f!, "id")),
                StringComparer.Ordinal);

            foreach (var coverage in GetArray(manifest, "fixtureClassCoverage", required: false))
            {
                var className = RequireString(coverage!, "class");
                var state = RequireString(coverage!, "state");
                _ = RequireString(coverage!, "notes");
                Assert.Contains(className, RequiredFixtureClasses);
                Assert.Contains(state, KnownCoverageStates);

                switch (state)
                {
                    case "verified":
                        var fixtureId = RequireString(coverage!, "fixtureId");
                        Assert.Contains(fixtureId, fixtureIds);
                        Assert.NotEmpty(GetStringArray(coverage!, "verifiedOperations"));
                        break;
                    case "blocked":
                        var reason = RequireString(coverage!, "reason");
                        Assert.Contains(reason, KnownReasons);
                        Assert.Null(coverage!["verifiedOperations"]);
                        break;
                    case "external-manual":
                        _ = RequireString(coverage!, "externalLane");
                        Assert.Null(coverage!["verifiedOperations"]);
                        break;
                }

                observed.Add(className);
            }
        }

        var missing = RequiredFixtureClasses.Where(c => !observed.Contains(c)).ToArray();
        Assert.True(
            missing.Length == 0,
            $"Required fixture classes missing from corpus: {string.Join(", ", missing)}");
    }

    [Fact]
    public void MalformedHwpxFixturesAreBlockedNotVerified()
    {
        var manifest = ReadJson("tests/fixtures/hwpx/manifest.json");
        var coverage = GetArray(manifest, "fixtureClassCoverage", required: false)
            .FirstOrDefault(c => RequireString(c!, "class") == "malformed-hwpx-package");
        Assert.NotNull(coverage);

        var state = RequireString(coverage!, "state");
        Assert.Equal("blocked", state);

        var reason = RequireString(coverage!, "reason");
        Assert.Equal(HwpCapabilityConstants.ReasonFixtureValidationFailed, reason);

        Assert.Null(coverage!["verifiedOperations"]);

        foreach (var path in ManifestPaths)
        {
            foreach (var fixture in GetArray(ReadJson(path), "fixtures"))
            {
                Assert.DoesNotContain(
                    "malformed-hwpx-package",
                    GetStringArray(fixture!, "classes"));
            }
        }
    }

    [Fact]
    public void ExternalManualFixturesDoNotCountAsVerified()
    {
        foreach (var path in ManifestPaths)
        {
            var manifest = ReadJson(path);
            foreach (var coverage in GetArray(manifest, "fixtureClassCoverage", required: false))
            {
                if (RequireString(coverage!, "state") != "external-manual") continue;
                _ = RequireString(coverage!, "externalLane");
                Assert.Null(coverage!["verifiedOperations"]);
                Assert.Null(coverage!["fixtureId"]);
            }
        }
    }

    [Fact]
    public void ManifestSchemaRejectsUnknownFormat()
    {
        var manifest = Clone(ReadJson("tests/fixtures/hwp/manifest.json"));
        manifest["format"] = "odt";

        AssertSchemaInvalid(LoadSchema(CompatibilityCorpusSchemaPath), manifest);
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

        AssertSchemaInvalid(LoadSchema(ExpectedCapabilitiesSchemaPath), capabilities);
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

    private static JsonSchema LoadSchema(string relativePath)
        => JsonSchema.FromText(File.ReadAllText(LocateRepoFile(relativePath)));

    private static void AssertSchemaValid(JsonSchema schema, JsonNode instance)
    {
        var results = schema.Evaluate(instance);
        Assert.True(results.IsValid, FormatSchemaErrors(results));
    }

    [Fact]
    public void Phase36ReleaseGateRequiresAllCorpusArtifacts()
    {
        string[] required =
        [
            "schemas/interfaces/compatibility-corpus.v1.schema.json",
            "schemas/interfaces/expected-capabilities.v1.schema.json",
            "tests/fixtures/common/expected-capabilities.json",
            "tests/fixtures/common/roundtrip-case.v1.schema.json",
            "tests/fixtures/common/roundtrip-cases.json",
            "tests/fixtures/common/visual-thresholds.json",
            "tests/fixtures/common/provider-compatibility.json",
            "tests/fixtures/hwp/manifest.json",
            "tests/fixtures/hwpx/manifest.json",
            "docs/qa/compatibility-corpus.md",
            "docs/qa/visual-diff-thresholds.md",
            "docs/qa/provider-compatibility-matrix.md",
            "docs/qa/phase-36-release-gate.md"
        ];

        foreach (var path in required)
        {
            _ = LocateRepoFile(path);
        }

        foreach (var manifestPath in ManifestPaths)
        {
            var manifest = ReadJson(manifestPath);
            Assert.NotNull(manifest["fixtureClassCoverage"]);
        }
    }

    [Fact]
    public void NoDocxParityLanguageBeforeScorecard()
    {
        string[] guardedDocs =
        [
            "docs/qa/compatibility-corpus.md",
            "docs/qa/visual-diff-thresholds.md",
            "docs/qa/provider-compatibility-matrix.md",
            "docs/qa/phase-36-release-gate.md"
        ];

        string[] forbiddenPhrases =
        [
            "DOCX parity",
            "docx parity",
            "DOCX 동등",
            "parity with DOCX"
        ];

        foreach (var doc in guardedDocs)
        {
            var lines = File.ReadAllLines(LocateRepoFile(doc));
            foreach (var phrase in forbiddenPhrases)
            {
                for (var i = 0; i < lines.Length; i++)
                {
                    if (!lines[i].Contains(phrase, StringComparison.OrdinalIgnoreCase)) continue;

                    var context = GetLocalClaimContext(lines, i);
                    if (HasScorecardGuard(context)) continue;

                    Assert.Fail(
                        $"{doc}:{i + 1} contains forbidden parity phrase '{phrase}' " +
                        "without local scorecard guard wording.");
                }
            }
        }
    }

    [Fact]
    public void BlockedOperationsRemainMachineReadable()
    {
        var capabilities = ReadJson(ExpectedCapabilitiesPath);
        var formats = capabilities["formats"]!.AsObject();

        var observedBlocked = 0;
        foreach (var (formatName, formatNode) in formats)
        {
            var ops = formatNode!["operations"]!.AsObject();
            foreach (var (opName, opNode) in ops)
            {
                var status = opNode!["status"]!.GetValue<string>();
                if (status != HwpCapabilityConstants.StatusUnsupported) continue;

                observedBlocked++;
                var reasonNode = opNode["reason"];
                Assert.True(
                    reasonNode is not null,
                    $"{formatName}.{opName} is unsupported but missing typed reason.");
                Assert.Contains(reasonNode!.GetValue<string>(), KnownReasons);
            }
        }

        var matrix = ReadJson("tests/fixtures/common/provider-compatibility.json");
        foreach (var row in matrix["rows"]!.AsArray())
        {
            var status = row!["status"]!.GetValue<string>();
            if (status != HwpCapabilityConstants.StatusUnsupported) continue;

            observedBlocked++;
            var reasonNode = row["blockedReason"];
            Assert.NotNull(reasonNode);
            Assert.Contains(reasonNode!.GetValue<string>(), KnownReasons);
        }

        Assert.True(observedBlocked > 0, "Expected at least one blocked operation across capability + matrix data.");
    }

    private static void AssertSchemaInvalid(JsonSchema schema, JsonNode instance)
    {
        var results = schema.Evaluate(instance);
        Assert.False(results.IsValid, "Schema unexpectedly accepted invalid corpus data.");
    }

    private static string FormatSchemaErrors(EvaluationResults results)
        => JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });

    private static string GetLocalClaimContext(string[] lines, int index)
    {
        var start = Math.Max(0, index - 3);
        var end = Math.Min(lines.Length - 1, index + 3);
        return string.Join('\n', lines[start..(end + 1)]);
    }

    private static bool HasScorecardGuard(string context)
    {
        string[] guardPhrases =
        [
            "forbidden claim",
            "forbidden until",
            "scorecard",
            "must not",
            "blocked until"
        ];

        return guardPhrases.Any(guard => context.Contains(guard, StringComparison.OrdinalIgnoreCase));
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
