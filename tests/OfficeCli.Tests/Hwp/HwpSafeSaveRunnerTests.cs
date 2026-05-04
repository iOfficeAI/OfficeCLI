using OfficeCli.Handlers.Hwp;
using OfficeCli.Handlers.Hwp.SafeSave;

namespace OfficeCli.Tests.Hwp;

[Collection("HwpBridgeEnvironment")]
public class HwpSafeSaveRunnerTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"officecli-safe-save-{Guid.NewGuid():N}");

    public HwpSafeSaveRunnerTests()
    {
        Directory.CreateDirectory(_root);
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    [Fact]
    public async Task OutputModeMovesVerifiedTempToOutput()
    {
        var input = CreateFile("input/source.hwp", "source");
        var output = Path.Combine(_root, "output/result.hwp");
        var transaction = await RunOutputAsync(input, output);

        Assert.True(transaction.Ok);
        Assert.True(transaction.Verified);
        Assert.True(File.Exists(output));
        Assert.Equal("edited", File.ReadAllText(output));
        Assert.Equal("output", transaction.Mode);
        Assert.Contains(transaction.Checks, check => check.Name == "temp-write" && check.Ok);
    }

    [Fact]
    public async Task OutputModeUsesOutputTargetDirectoryForTemp()
    {
        var input = CreateFile("input/source.hwp", "source");
        var outputDir = Path.Combine(_root, "other-volume-like-output");
        var output = Path.Combine(outputDir, "result.hwp");

        var transaction = await RunOutputAsync(input, output);

        Assert.NotNull(transaction.TempPath);
        Assert.Equal(outputDir, Path.GetDirectoryName(transaction.TempPath));
    }

    [Fact]
    public async Task MissingRequiredCheckFailsTransaction()
    {
        var input = CreateFile("input/source.hwp", "source");
        var output = Path.Combine(_root, "output/result.hwp");
        var runner = new SafeSaveRunner();

        var transaction = await runner.RunAsync(
            Options(input, output, SafeSavePolicy.OutputMode("temp-write", "semantic-delta")),
            temp =>
            {
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks([])),
            CancellationToken.None);

        Assert.False(transaction.Ok);
        Assert.False(File.Exists(output));
        Assert.Contains(transaction.Checks, check => check.Name == "required-checks" && !check.Ok);
        Assert.Equal("source", File.ReadAllText(input));
    }

    [Fact]
    public async Task SameInputOutputPathFailsBeforeWritingTemp()
    {
        var input = CreateFile("input/source.hwp", "source");
        var runner = new SafeSaveRunner();
        var writeCalled = false;

        var transaction = await runner.RunAsync(
            Options(input, input, SafeSavePolicy.OutputMode("temp-write")),
            temp =>
            {
                writeCalled = true;
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks([])),
            CancellationToken.None);

        Assert.False(transaction.Ok);
        Assert.False(writeCalled);
        Assert.Equal("source", File.ReadAllText(input));
        Assert.Contains(transaction.Checks, check => check.Name == "same-path-output" && !check.Ok);
    }

    [Fact]
    public async Task CaseVariantSamePathFailsOnCaseInsensitivePlatforms()
    {
        if (!OperatingSystem.IsMacOS() && !OperatingSystem.IsWindows()) return;

        var input = CreateFile("input/source.hwp", "source");
        var variant = input.ToUpperInvariant();
        var runner = new SafeSaveRunner();
        var writeCalled = false;

        var transaction = await runner.RunAsync(
            Options(input, variant, SafeSavePolicy.OutputMode("temp-write")),
            temp =>
            {
                writeCalled = true;
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks([])),
            CancellationToken.None);

        Assert.False(transaction.Ok);
        Assert.False(writeCalled);
        Assert.Equal("source", File.ReadAllText(input));
        Assert.Contains(transaction.Checks, check => check.Name == "same-path-output" && !check.Ok);
    }

    [Fact]
    public async Task ValidationFailureLeavesInputAndExistingOutputUntouched()
    {
        var input = CreateFile("input/source.hwp", "source");
        var output = CreateFile("output/result.hwp", "previous");
        var runner = new SafeSaveRunner();

        var transaction = await runner.RunAsync(
            Options(input, output, SafeSavePolicy.OutputMode("temp-write", "provider-readback")),
            temp =>
            {
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks(
            [
                new SafeSaveCheck("provider-readback", false, "error", "readback failed")
            ])),
            CancellationToken.None);

        Assert.False(transaction.Ok);
        Assert.Equal("source", File.ReadAllText(input));
        Assert.Equal("previous", File.ReadAllText(output));
    }

    [Fact]
    public async Task InPlaceReturnsNotReadyAndDoesNotReplaceSource()
    {
        var input = CreateFile("input/source.hwp", "source");
        var runner = new SafeSaveRunner();

        var transaction = await runner.RunAsync(
            Options(input, input, SafeSavePolicy.OutputMode("temp-write")) with { InPlace = true },
            temp =>
            {
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks([])),
            CancellationToken.None);

        Assert.False(transaction.Ok);
        Assert.False(transaction.Verified);
        Assert.Equal("source", File.ReadAllText(input));
        Assert.Contains(transaction.Checks, check => check.Name == "in-place-not-ready" && !check.Ok);
    }

    private async Task<SafeSaveTransaction> RunOutputAsync(string input, string output)
    {
        var runner = new SafeSaveRunner();
        return await runner.RunAsync(
            Options(input, output, SafeSavePolicy.OutputMode("temp-write")),
            temp =>
            {
                File.WriteAllText(temp, "edited");
                return Task.CompletedTask;
            },
            _ => Task.FromResult(SafeSaveValidationResult.FromChecks([])),
            CancellationToken.None);
    }

    private SafeSaveOptions Options(string input, string output, SafeSavePolicy policy) => new(
        input,
        output,
        InPlace: false,
        Backup: false,
        Verify: true,
        HwpCapabilityConstants.OperationReplaceText,
        HwpCapabilityConstants.FormatHwp,
        policy);

    private string CreateFile(string relativePath, string content)
    {
        var path = Path.Combine(_root, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
        return path;
    }
}
