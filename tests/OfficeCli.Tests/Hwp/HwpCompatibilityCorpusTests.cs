using System.Security.Cryptography;
using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public sealed class HwpCompatibilityCorpusTests
{
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
    public void ExpectedCapabilities_ReferenceKnownOperationsAndEvidence()
    {
        var manifest = ReadJson("tests/fixtures/common/expected-capabilities.json");
        var knownOperations = new HashSet<string>
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

        foreach (var format in manifest["formats"]!.AsObject())
        {
            Assert.True(
                new[] { HwpCapabilityConstants.FormatHwp, HwpCapabilityConstants.FormatHwpx }.Contains(format.Key),
                $"Unknown corpus format '{format.Key}'.");
            var operations = format.Value!["operations"]!.AsObject();
            Assert.NotEmpty(operations);
            foreach (var operation in operations)
            {
                Assert.Contains(operation.Key, knownOperations);
                var status = operation.Value!["status"]!.GetValue<string>();
                Assert.True(
                    new[] { "experimental", "roundtrip-verified", "unsupported" }.Contains(status),
                    $"Unknown status '{status}' for {format.Key}.{operation.Key}.");
                foreach (var evidence in operation.Value["evidence"]!.AsArray())
                    _ = LocateRepoFile(evidence!.GetValue<string>());
            }
        }
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

    private static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
}
