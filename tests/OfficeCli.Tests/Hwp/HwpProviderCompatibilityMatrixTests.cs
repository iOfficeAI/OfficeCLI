// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public sealed class HwpProviderCompatibilityMatrixTests
{
    private const string MatrixPath = "tests/fixtures/common/provider-compatibility.json";
    private const string ExpectedCapabilitiesPath = "tests/fixtures/common/expected-capabilities.json";

    private static readonly string[] MatrixProviders =
    [
        HwpCapabilityConstants.EngineCustom,
        HwpCapabilityConstants.EngineRhwpBridge
    ];

    private static readonly HashSet<string> AllowedStatuses = new(StringComparer.Ordinal)
    {
        HwpCapabilityConstants.StatusUnsupported,
        HwpCapabilityConstants.StatusExperimental,
        "fixture-backed",
        HwpCapabilityConstants.StatusRoundTripVerified,
        "external-manual"
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
    public void HwpxCustomRemainsDefault()
    {
        var hwpxDefaultRows = LoadRows()
            .Where(r =>
                r["format"]!.GetValue<string>() == HwpCapabilityConstants.FormatHwpx
                && r["defaultProvider"]!.GetValue<bool>())
            .ToList();

        Assert.NotEmpty(hwpxDefaultRows);

        foreach (var row in hwpxDefaultRows)
        {
            Assert.Equal(
                HwpCapabilityConstants.EngineCustom,
                row["provider"]!.GetValue<string>());
        }
    }

    [Fact]
    public void RhwpPromotionRequiresEvidenceParity()
    {
        foreach (var row in LoadRows())
        {
            var provider = row["provider"]!.GetValue<string>();
            if (provider != HwpCapabilityConstants.EngineRhwpBridge) continue;

            var status = row["status"]!.GetValue<string>();
            if (status is "fixture-backed" or "roundtrip-verified")
            {
                var evidence = row["evidence"]?.AsArray();
                Assert.True(
                    evidence is { Count: > 0 },
                    $"rhwp-bridge row promoted to '{status}' without evidence: " +
                    $"{row["format"]!.GetValue<string>()}.{row["operation"]!.GetValue<string>()}");
            }

            if (row["format"]!.GetValue<string>() == HwpCapabilityConstants.FormatHwpx)
            {
                Assert.False(
                    row["defaultProvider"]!.GetValue<bool>(),
                    "HWPX rhwp-bridge rows must not be the default provider.");
            }
        }
    }

    [Fact]
    public void HancomLaneIsOptionalNotCiRequired()
    {
        var allowed = new HashSet<string>(StringComparer.Ordinal) { "optional", "external-manual" };
        foreach (var row in LoadRows())
        {
            var lane = row["hancomLane"]!.GetValue<string>();
            Assert.Contains(lane, allowed);
            Assert.NotEqual("required", lane);
        }
    }

    [Fact]
    public void BlockedProviderRowsHaveTypedReasons()
    {
        foreach (var row in LoadRows())
        {
            var status = row["status"]!.GetValue<string>();
            Assert.Contains(status, AllowedStatuses);

            if (status != HwpCapabilityConstants.StatusUnsupported) continue;

            var reasonNode = row["blockedReason"];
            Assert.NotNull(reasonNode);
            var reason = reasonNode!.GetValue<string>();
            Assert.Contains(reason, KnownReasons);
        }
    }

    [Fact]
    public void RowsAreUnique()
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in LoadRows())
        {
            var key = MatrixKey(row);
            Assert.True(seen.Add(key), $"Duplicate provider matrix row: {key}");
        }
    }

    [Fact]
    public void MatrixCoversExpectedCapabilityProviderPairs()
    {
        var rowsByKey = LoadRows().ToDictionary(MatrixKey, StringComparer.Ordinal);
        var capabilities = JsonNode.Parse(File.ReadAllText(LocateRepoFile(ExpectedCapabilitiesPath)))!;
        var formats = capabilities["formats"]!.AsObject();

        foreach (var (format, formatNode) in formats)
        {
            var defaultEngine = formatNode!["defaultEngine"]!.GetValue<string>();
            var operations = formatNode["operations"]!.AsObject();

            foreach (var (operation, _) in operations)
            {
                var defaultRows = 0;
                foreach (var provider in MatrixProviders)
                {
                    var key = MatrixKey(format, operation, provider);
                    Assert.True(rowsByKey.ContainsKey(key), $"Missing provider matrix row: {key}");

                    if (!rowsByKey[key]["defaultProvider"]!.GetValue<bool>()) continue;

                    defaultRows++;
                    Assert.Equal(
                        defaultEngine,
                        rowsByKey[key]["provider"]!.GetValue<string>());
                }

                Assert.Equal(1, defaultRows);
            }
        }
    }

    private static IEnumerable<JsonNode> LoadRows()
    {
        var doc = JsonNode.Parse(File.ReadAllText(LocateRepoFile(MatrixPath)))!;
        foreach (var row in doc["rows"]!.AsArray())
        {
            Assert.NotNull(row);
            yield return row!;
        }
    }

    private static string MatrixKey(JsonNode row)
        => MatrixKey(
            row["format"]!.GetValue<string>(),
            row["operation"]!.GetValue<string>(),
            row["provider"]!.GetValue<string>());

    private static string MatrixKey(string? format, string? operation, string provider)
        => $"{format}.{operation}.{provider}";

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
