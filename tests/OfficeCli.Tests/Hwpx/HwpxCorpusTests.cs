// Plan 86: Corpus Smoke Validation
// Runs open → view text → validate on real HWPX samples.
// Corpus directory is optional (local-only, .gitignore) — tests skip if absent.

using OfficeCli.Handlers;

namespace OfficeCli.Tests.Hwpx;

public class HwpxCorpusTests
{
    // Corpus: jakal-hwpx samples (copied or symlinked locally)
    private static readonly string CorpusDir = Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Fixtures", "corpus");

    // Fallback: reference repo samples (read-only, no save/reopen)
    private static readonly string ReferenceCorpusDir = Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "..", "..", "..",
        "devlog", "_plan", "office", "hwp", "reference", "repos",
        "jakal-hwpx", "examples", "samples", "hwpx");

    private static string? GetCorpusDir()
    {
        if (Directory.Exists(CorpusDir)) return CorpusDir;
        if (Directory.Exists(ReferenceCorpusDir)) return ReferenceCorpusDir;
        return null;
    }

    public static IEnumerable<object[]> CorpusFiles()
    {
        var dir = GetCorpusDir();
        if (dir == null) yield break;
        foreach (var f in Directory.GetFiles(dir, "*.hwpx"))
            yield return [Path.GetFileName(f), f];
    }

    [Theory]
    [MemberData(nameof(CorpusFiles))]
    public void Smoke_Open(string name, string path)
    {
        // Just open — should not throw
        using var handler = new HwpxHandler(path, editable: false);
        Assert.NotNull(handler);
    }

    [Theory]
    [MemberData(nameof(CorpusFiles))]
    public void Smoke_ViewText(string name, string path)
    {
        using var handler = new HwpxHandler(path, editable: false);
        var text = handler.ViewAsText();
        // Real documents should have some text
        Assert.False(string.IsNullOrEmpty(text), $"{name}: ViewAsText returned empty");
    }

    [Theory]
    [MemberData(nameof(CorpusFiles))]
    public void Smoke_ViewMarkdown(string name, string path)
    {
        using var handler = new HwpxHandler(path, editable: false);
        var md = handler.ViewAsMarkdown();
        Assert.NotNull(md);
    }

    [Theory]
    [MemberData(nameof(CorpusFiles))]
    public void Smoke_ViewForms(string name, string path)
    {
        using var handler = new HwpxHandler(path, editable: false);
        var forms = handler.ViewAsForms(auto: true);
        Assert.NotNull(forms);
    }

    [Theory]
    [MemberData(nameof(CorpusFiles))]
    public void Smoke_Validate(string name, string path)
    {
        using var handler = new HwpxHandler(path, editable: false);
        // Should not crash — warnings/errors are OK
        var errors = handler.Validate();
        Assert.NotNull(errors);
    }
}
