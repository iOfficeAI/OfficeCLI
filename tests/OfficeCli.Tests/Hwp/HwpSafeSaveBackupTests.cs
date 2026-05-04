using OfficeCli.Handlers.Hwp;
using OfficeCli.Handlers.Hwp.SafeSave;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public sealed class HwpSafeSaveBackupTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"officecli-safe-save-backup-{Guid.NewGuid():N}");

    public HwpSafeSaveBackupTests()
    {
        Directory.CreateDirectory(_root);
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    [Fact]
    public async Task InPlaceCreatesBackupBeforeReplacingSource()
    {
        var input = CreateFile("source.hwp", "source");

        var transaction = await RunInPlaceAsync(new SafeSaveRunner(), input);

        Assert.True(transaction.Ok);
        Assert.True(transaction.Verified);
        Assert.Equal("edited", File.ReadAllText(input));
        Assert.NotNull(transaction.BackupPath);
        Assert.Equal("source", File.ReadAllText(transaction.BackupPath!));
        Assert.NotNull(transaction.ManifestPath);
        Assert.True(File.Exists(transaction.ManifestPath));
        Assert.Contains(transaction.Checks, check => check.Name == "backup-created" && check.Ok);
        Assert.Contains(transaction.Checks, check => check.Name == "manifest-write" && check.Ok);
    }

    [Fact]
    public async Task InPlaceRefusesReplaceWhenManifestWriteFails()
    {
        var input = CreateFile("source.hwp", "source");
        var runner = new SafeSaveRunner(new ThrowingManifestWriter(_root));

        var transaction = await RunInPlaceAsync(runner, input);

        Assert.False(transaction.Ok);
        Assert.Equal("source", File.ReadAllText(input));
        Assert.NotNull(transaction.BackupPath);
        Assert.Equal("source", File.ReadAllText(transaction.BackupPath!));
        Assert.Contains(transaction.Checks, check => check.Name == "manifest-write" && !check.Ok);
    }

    private static async Task<SafeSaveTransaction> RunInPlaceAsync(SafeSaveRunner runner, string input)
        => await runner.RunAsync(
            Options(input),
            temp =>
            {
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks(
            [
                new SafeSaveCheck("provider-readback", true, "info")
            ])),
            CancellationToken.None);

    private static SafeSaveOptions Options(string input) => new(
        input,
        input,
        InPlace: true,
        Backup: true,
        Verify: true,
        HwpCapabilityConstants.OperationReplaceText,
        HwpCapabilityConstants.FormatHwp,
        SafeSavePolicy.OutputMode("temp-write", "provider-readback"));

    private string CreateFile(string relativePath, string content)
    {
        var path = Path.Combine(_root, relativePath);
        File.WriteAllText(path, content);
        return path;
    }

    private sealed class ThrowingManifestWriter(string root) : ISafeSaveManifestWriter
    {
        public string BuildManifestPath(SafeSaveOptions options, DateTimeOffset timestamp)
            => Path.Combine(root, "blocked.officecli-transaction.json");

        public void Write(SafeSaveTransaction transaction)
            => throw new IOException("manifest blocked");
    }
}
