using System.Text.Json.Nodes;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli.Tests.Hwp;

public sealed class HwpRhwpSurfaceParityTests
{
    [Fact]
    public void Matrix_ClassifiesEveryAuditedRhwpCliAndBridgeSurface()
    {
        var root = LoadMatrix();
        var rows = root["rows"]!.AsArray().Select(node => node!.AsObject()).ToArray();
        var ids = rows.Select(row => RequiredString(row, "id")).ToArray();

        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(18, root["cliCommandCounts"]!["pinned"]!.GetValue<int>());
        Assert.Equal(19, root["cliCommandCounts"]!["latest"]!.GetValue<int>());

        AssertRequiredIds(ids, [
            "cli.convert",
            "cli.dump",
            "cli.info",
            "cli.thumbnail",
            "cli.build-from-ingest",
            "bridge.insert-text",
            "bridge.get-cell-text",
            "bridge.scan-cells",
            "api.convert_to_editable_native",
            "api.render_page_html",
            "api.begin_batch",
            "api.insert_text_native"
        ]);
    }

    [Fact]
    public void Matrix_HasGateMetadataForP0AndP1Rows()
    {
        var root = LoadMatrix();
        var rows = root["rows"]!.AsArray().Select(node => node!.AsObject()).ToArray();
        var allowedClasses = root["allowedClasses"]!.AsArray()
            .Select(node => node!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);
        var allowedStatuses = root["allowedStatuses"]!.AsArray()
            .Select(node => node!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);

        foreach (var row in rows)
        {
            var id = RequiredString(row, "id");
            var classification = RequiredString(row, "class");
            var status = RequiredString(row, "officecliStatus");
            Assert.Contains(classification, allowedClasses);
            Assert.Contains(status, allowedStatuses);
            Assert.False(string.IsNullOrWhiteSpace(RequiredString(row, "rhwp")), id);

            if (classification is "P0" or "P1")
            {
                Assert.True(
                    row.ContainsKey("capability") || status is "blocked" or "dev-only",
                    $"{id} must have a capability name or be explicitly blocked/dev-only.");
                Assert.True(
                    row.ContainsKey("officecliSurface") || row.ContainsKey("phase") || row.ContainsKey("blockedReason"),
                    $"{id} must have an OfficeCLI surface, phase, or blocked reason.");
            }
        }
    }

    [Fact]
    public void Matrix_TracksAllCurrentBridgeCommands()
    {
        var root = LoadMatrix();
        var bridgeCommands = root["rows"]!.AsArray()
            .Select(node => node!.AsObject())
            .Where(row => RequiredString(row, "kind") == "officecli-bridge")
            .Select(row => RequiredString(row, "rhwp"))
            .ToHashSet(StringComparer.Ordinal);

        Assert.Equal(new[]
        {
            "convert-to-editable",
            "create-blank",
            "diagnostics",
            "document-info",
            "dump-controls",
            "dump-pages",
            "export-pdf",
            "export-markdown",
            "get-cell-text",
            "get-field",
            "insert-text",
            "list-fields",
            "native-op",
            "read-text",
            "render-png",
            "render-svg",
            "replace-text",
            "save-as-hwp",
            "scan-cells",
            "set-cell-text",
            "set-field",
            "thumbnail"
        }.Order(StringComparer.Ordinal), bridgeCommands.Order(StringComparer.Ordinal));
    }

    [Fact]
    public void Matrix_OperationCapabilitiesExistForWiredRows()
    {
        var capabilityNames = typeof(HwpCapabilityConstants)
            .GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(field => field.Name.StartsWith("Operation", StringComparison.Ordinal))
            .Select(field => (string)field.GetValue(null)!)
            .ToHashSet(StringComparer.Ordinal);
        var root = LoadMatrix();
        var wiredRows = root["rows"]!.AsArray()
            .Select(node => node!.AsObject())
            .Where(row => RequiredString(row, "officecliStatus") == "wired" && row.ContainsKey("capability"));

        foreach (var row in wiredRows)
        {
            var capability = RequiredString(row, "capability");
            Assert.Contains(capability, capabilityNames);
        }
    }

    private static JsonObject LoadMatrix()
        => JsonNode.Parse(File.ReadAllText(LocateRepoFile("tests/fixtures/common/rhwp-surface-parity.json")))!
            .AsObject();

    private static void AssertRequiredIds(IReadOnlyCollection<string> ids, IEnumerable<string> required)
    {
        foreach (var id in required)
            Assert.Contains(id, ids);
    }

    private static string RequiredString(JsonObject row, string key)
    {
        Assert.True(row.ContainsKey(key), $"Missing required key '{key}' in row: {row}");
        return row[key]!.GetValue<string>();
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
