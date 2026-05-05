// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;
using Json.Schema;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public sealed class HwpRoundTripCorpusTests
{
    private const string CasesPath = "tests/fixtures/common/roundtrip-cases.json";
    private const string CasesSchemaPath = "tests/fixtures/common/roundtrip-case.v1.schema.json";
    private const string ManifestSchemaPath = "schemas/interfaces/compatibility-corpus.v1.schema.json";

    private static readonly string[] ManifestPaths =
    [
        "tests/fixtures/hwp/manifest.json",
        "tests/fixtures/hwpx/manifest.json"
    ];

    private static readonly HashSet<string> MutationOperations = new(StringComparer.Ordinal)
    {
        HwpCapabilityConstants.OperationFillField,
        HwpCapabilityConstants.OperationReplaceText,
        HwpCapabilityConstants.OperationSetTableCell,
        HwpCapabilityConstants.OperationSaveOriginal,
        HwpCapabilityConstants.OperationSaveAsHwp
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

    [Fact]
    public void CasesConformToSchema()
    {
        var schema = LoadSchema(CasesSchemaPath);
        var catalog = ReadJson(CasesPath);

        var results = schema.Evaluate(catalog);
        Assert.True(results.IsValid, FormatErrors(results));

        var cases = catalog["cases"]!.AsArray();
        Assert.NotEmpty(cases);

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var fixtureIds = LoadAllFixtureIds();

        foreach (var item in cases)
        {
            var caseId = RequireString(item!, "id");
            Assert.True(seen.Add(caseId), $"Duplicate round-trip case id: {caseId}");

            var fixtureId = RequireString(item!, "fixtureId");
            Assert.Contains(fixtureId, fixtureIds);

            var inputPath = RequireString(item!, "inputPath");
            _ = LocateRepoFile(inputPath);
        }
    }

    [Fact]
    public void MutationCasesLeaveSourceUnchanged()
    {
        foreach (var item in LoadCases())
        {
            var operation = RequireString(item, "operation");
            if (!MutationOperations.Contains(operation)) continue;

            var requiredChecks = GetStringArray(item, "requiredChecks");
            Assert.Contains("source-unchanged", requiredChecks);

            var outputMode = RequireString(item, "outputMode");
            Assert.NotEqual("in-place", outputMode);
        }
    }

    [Fact]
    public void BlockedCasesReturnTypedErrors()
    {
        var observedHwpxBlocked = false;

        foreach (var item in LoadCases())
        {
            var expected = item["expected"]!.AsObject();
            var status = RequireString(expected, "status");
            if (status != "blocked") continue;

            var error = expected["error"]?.AsObject();
            Assert.NotNull(error);
            var code = RequireString(error!, "code");
            Assert.Contains(code, KnownReasons);

            Assert.Contains("typed-error-if-blocked", GetStringArray(item, "requiredChecks"));

            observedHwpxBlocked |= RequireString(item, "format") == HwpCapabilityConstants.FormatHwpx
                && RequireString(item, "operation") == HwpCapabilityConstants.OperationSetTableCell
                && code == HwpCapabilityConstants.ReasonRoundTripUnverified;
        }

        Assert.True(
            observedHwpxBlocked,
            "HWPX set_table_cell must be present as a blocked case with reason roundtrip_unverified.");
    }

    [Fact]
    public void SuccessCasesPassProviderReadback()
    {
        var observedSuccess = 0;
        foreach (var item in LoadCases())
        {
            var status = RequireString(item["expected"]!, "status");
            if (status != "success") continue;

            var requiredChecks = GetStringArray(item, "requiredChecks");
            Assert.Contains("provider-readback", requiredChecks);
            observedSuccess++;
        }

        Assert.True(observedSuccess > 0, "At least one success-status case is required.");
    }

    [Fact]
    public void RealRoundTripExecutionIsOptIn()
    {
        var realRhwp = Environment.GetEnvironmentVariable("OFFICECLI_REAL_RHWP_BIN");
        if (string.IsNullOrWhiteSpace(realRhwp))
        {
            Assert.True(
                LoadCases().Any(),
                "Round-trip catalog must remain populated even when OFFICECLI_REAL_RHWP_BIN is unset.");
            return;
        }

        Assert.True(
            File.Exists(realRhwp),
            $"OFFICECLI_REAL_RHWP_BIN points at a missing path: {realRhwp}");
    }

    private static IEnumerable<JsonNode> LoadCases()
    {
        var catalog = ReadJson(CasesPath);
        foreach (var item in catalog["cases"]!.AsArray())
        {
            Assert.NotNull(item);
            yield return item!;
        }
    }

    private static HashSet<string> LoadAllFixtureIds()
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var path in ManifestPaths)
        {
            foreach (var fixture in ReadJson(path)["fixtures"]!.AsArray())
            {
                ids.Add(RequireString(fixture!, "id"));
            }
        }
        return ids;
    }

    private static JsonSchema LoadSchema(string relativePath)
        => JsonSchema.FromText(File.ReadAllText(LocateRepoFile(relativePath)));

    private static JsonNode ReadJson(string relativePath)
        => JsonNode.Parse(File.ReadAllText(LocateRepoFile(relativePath)))!;

    private static string FormatErrors(EvaluationResults results)
        => JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });

    private static string[] GetStringArray(JsonNode node, string propertyName)
    {
        var array = node[propertyName]?.AsArray()
            ?? throw new InvalidOperationException($"Missing required array '{propertyName}'.");
        return array.Select(item => RequireString(item!)).ToArray();
    }

    private static string RequireString(JsonNode node, string propertyName)
        => RequireString(node[propertyName] ?? throw new InvalidOperationException($"Missing required string '{propertyName}'."));

    private static string RequireString(JsonNode node)
    {
        var value = node.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value))
            throw new InvalidOperationException("Expected a non-empty string.");
        return value;
    }

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
}
