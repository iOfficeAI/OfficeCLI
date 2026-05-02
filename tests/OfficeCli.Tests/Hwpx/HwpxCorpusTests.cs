// Plan 86: Corpus Smoke Validation
// Runs open → view text → validate on real HWPX samples.
// Corpus directory is optional (local-only, .gitignore) — tests no-op if absent.

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

    private static IEnumerable<(string Name, string Path)> GetFiles()
    {
        var dir = GetCorpusDir();
        if (dir == null) yield break;
        foreach (var f in Directory.GetFiles(dir, "*.hwpx"))
            yield return (Path.GetFileName(f), f);
    }

    [Fact]
    public void Smoke_Open()
    {
        foreach (var (name, path) in GetFiles())
        {
            using var handler = new HwpxHandler(path, editable: false);
            Assert.NotNull(handler);
        }
    }

    [Fact]
    public void Smoke_ViewText()
    {
        foreach (var (name, path) in GetFiles())
        {
            using var handler = new HwpxHandler(path, editable: false);
            var text = handler.ViewAsText();
            Assert.False(string.IsNullOrEmpty(text), $"{name}: ViewAsText returned empty");
        }
    }

    [Fact]
    public void Smoke_ViewMarkdown()
    {
        foreach (var (name, path) in GetFiles())
        {
            using var handler = new HwpxHandler(path, editable: false);
            var md = handler.ViewAsMarkdown();
            Assert.NotNull(md);
        }
    }

    [Fact]
    public void Smoke_ViewForms()
    {
        foreach (var (name, path) in GetFiles())
        {
            using var handler = new HwpxHandler(path, editable: false);
            var forms = handler.ViewAsForms(auto: true);
            Assert.NotNull(forms);
        }
    }

    [Fact]
    public void Smoke_Validate()
    {
        foreach (var (name, path) in GetFiles())
        {
            using var handler = new HwpxHandler(path, editable: false);
            var errors = handler.Validate();
            Assert.NotNull(errors);
        }
    }
}
