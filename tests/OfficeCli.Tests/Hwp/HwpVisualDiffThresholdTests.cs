// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public sealed class HwpVisualDiffThresholdTests
{
    private const string ThresholdsPath = "tests/fixtures/common/visual-thresholds.json";
    private const string ExpectedCapabilitiesPath = "tests/fixtures/common/expected-capabilities.json";

    [Fact]
    public void PageCountMismatchIsHardFail()
    {
        var thresholds = ReadJson(ThresholdsPath);
        var hardFail = thresholds["hardFail"]!.AsArray()
            .Select(n => n!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("page-count-mismatch", hardFail);
        Assert.Contains("missing-expected-svg-page", hardFail);

        var thresholded = thresholds["thresholdedFail"]!.AsArray();
        foreach (var row in thresholded)
        {
            var id = row!["id"]!.GetValue<string>();
            Assert.DoesNotContain("page-count", id);
        }
    }

    [Fact]
    public void MissingRenderEvidenceFailsVisualClaim()
    {
        var thresholds = ReadJson(ThresholdsPath);
        var visualValidated = thresholds["visualValidatedOperations"]!.AsArray()
            .Select(n => n!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);

        var hardFail = thresholds["hardFail"]!.AsArray()
            .Select(n => n!.GetValue<string>())
            .ToList();
        Assert.Contains(
            "missing-render-evidence-for-visual-validated-operation",
            hardFail);

        var capabilities = ReadJson(ExpectedCapabilitiesPath);
        var formats = capabilities["formats"]!.AsObject();

        foreach (var (formatName, formatNode) in formats)
        {
            var ops = formatNode!["operations"]!.AsObject();
            foreach (var (opName, opNode) in ops)
            {
                if (!visualValidated.Contains(opName)) continue;
                var status = opNode!["status"]!.GetValue<string>();
                if (status == HwpCapabilityConstants.StatusUnsupported) continue;

                var evidence = opNode["evidence"]?.AsArray();
                Assert.True(
                    evidence is { Count: > 0 },
                    $"{formatName}.{opName} is visual-validated with status '{status}' but has no render evidence.");
            }
        }
    }

    [Fact]
    public void TextOnlyMutationUsesDeclaredThreshold()
    {
        var thresholds = ReadJson(ThresholdsPath);
        var rows = thresholds["thresholdedFail"]!.AsArray();

        JsonNode? layoutDrift = null;
        foreach (var row in rows)
        {
            if (row!["id"]!.GetValue<string>() == "text-only-mutation-layout-drift")
            {
                layoutDrift = row;
                break;
            }
        }

        Assert.NotNull(layoutDrift);
        Assert.Equal("layout-drift-fraction", layoutDrift!["metric"]!.GetValue<string>());

        var maxFraction = layoutDrift["maxFraction"]!.GetValue<double>();
        Assert.InRange(maxFraction, 0.0, 0.05);

        var appliesTo = layoutDrift["appliesTo"]!.AsArray()
            .Select(n => n!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);
        Assert.Contains(HwpCapabilityConstants.OperationReplaceText, appliesTo);
        Assert.Contains(HwpCapabilityConstants.OperationFillField, appliesTo);
    }

    private static JsonNode ReadJson(string relativePath)
        => JsonNode.Parse(File.ReadAllText(LocateRepoFile(relativePath)))!;

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
