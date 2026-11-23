using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.MobileCore;

/// <summary>
/// In-process, mobile-safe facade over OfficeCLI's OOXML document engine.
/// It deliberately exposes no shell, process, pipe, plugin, installer, watch,
/// browser, arbitrary output path, or raw-package operations.
/// </summary>
public sealed class MobileDocumentSession : IDisposable
{
    private readonly string _documentPath;
    private readonly IDocumentHandler _handler;
    private bool _disposed;

    private MobileDocumentSession(string documentPath, IDocumentHandler handler)
    {
        _documentPath = documentPath;
        _handler = handler;
    }

    public string DocumentPath => _documentPath;

    public static MobileDocumentSession Open(string appPrivateDocumentPath)
    {
        var fullPath = Path.GetFullPath(appPrivateDocumentPath);
        var extension = Path.GetExtension(fullPath);
        if (extension is not (".docx" or ".xlsx" or ".pptx"))
            throw new NotSupportedException("Only DOCX, XLSX, and PPTX are supported on mobile.");
        return new MobileDocumentSession(fullPath, DocumentHandlerFactory.Open(fullPath, editable: true));
    }

    public static MobileDocumentSession Create(string appPrivateDocumentPath)
    {
        var fullPath = Path.GetFullPath(appPrivateDocumentPath);
        var extension = Path.GetExtension(fullPath);
        if (extension is not (".docx" or ".xlsx" or ".pptx"))
            throw new NotSupportedException("Only DOCX, XLSX, and PPTX are supported on mobile.");
        OfficeCli.BlankDocCreator.Create(fullPath);
        return Open(fullPath);
    }

    public string RenderHtml()
    {
        ThrowIfDisposed();
        return _handler switch
        {
            WordHandler word => word.ViewAsHtml(),
            ExcelHandler excel => excel.ViewAsHtml(),
            PowerPointHandler powerPoint => powerPoint.ViewAsHtml(),
            _ => throw new NotSupportedException("This document type has no mobile HTML renderer.")
        };
    }

    public JsonNode Outline()
    {
        ThrowIfDisposed();
        return _handler.ViewAsOutlineJson();
    }

    public JsonNode Get(string selector, int depth = 1)
    {
        ThrowIfDisposed();
        ValidateSelector(selector);
        if (depth is < 0 or > 10) throw new ArgumentOutOfRangeException(nameof(depth));
        return JsonSerializer.SerializeToNode(_handler.Get(selector, depth))!;
    }

    public IReadOnlyList<JsonNode> Query(string selector)
    {
        ThrowIfDisposed();
        ValidateSelector(selector);
        return _handler.Query(selector).Select(node => JsonSerializer.SerializeToNode(node)!).ToArray();
    }

    public MobileCommandResult Execute(MobileOfficeCommand command)
        => ExecuteCore(command, save: true);

    public IReadOnlyList<MobileCommandResult> ExecuteBatch(IReadOnlyList<MobileOfficeCommand> commands)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(commands);
        if (commands.Count is < 1 or > 100) throw new ArgumentException("A batch must contain between 1 and 100 commands.");
        var results = new List<MobileCommandResult>(commands.Count);
        foreach (var command in commands) results.Add(ExecuteCore(command, save: false));
        _handler.Save();
        return results;
    }

    private MobileCommandResult ExecuteCore(MobileOfficeCommand command, bool save)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(command);
        ValidateSelector(command.Path);
        var properties = command.Properties ?? new Dictionary<string, string>();
        if (properties.Count > 100) throw new ArgumentException("A command may contain at most 100 properties.");

        string? message;
        IReadOnlyList<string> unsupported = [];
        switch (command.Operation)
        {
            case MobileOperation.Set:
                unsupported = _handler.Set(command.Path, properties);
                message = "Element updated.";
                break;
            case MobileOperation.Add:
                if (string.IsNullOrWhiteSpace(command.Type)) throw new ArgumentException("Add requires an element type.");
                message = _handler.Add(command.Path, command.Type, ToPosition(command), properties);
                break;
            case MobileOperation.Remove:
                message = _handler.Remove(command.Path, properties);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(command.Operation));
        }

        if (save) _handler.Save();
        return new MobileCommandResult(true, message, unsupported);
    }

    public void Save()
    {
        ThrowIfDisposed();
        _handler.Save();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _handler.Dispose();
        _disposed = true;
    }

    private static InsertPosition? ToPosition(MobileOfficeCommand command)
    {
        var count = (command.Index.HasValue ? 1 : 0) + (command.Before is not null ? 1 : 0) + (command.After is not null ? 1 : 0);
        if (count > 1) throw new ArgumentException("Only one insertion position may be specified.");
        if (command.Index.HasValue) return InsertPosition.AtIndex(command.Index.Value);
        if (command.Before is not null) return InsertPosition.BeforeElement(command.Before);
        if (command.After is not null) return InsertPosition.AfterElement(command.After);
        return null;
    }

    private static void ValidateSelector(string selector)
    {
        if (string.IsNullOrWhiteSpace(selector) || !selector.StartsWith('/'))
            throw new ArgumentException("An OfficeCLI document selector beginning with '/' is required.");
        if (selector.Length > 2048) throw new ArgumentException("The selector is too long.");
        if (selector.Contains("..", StringComparison.Ordinal) || selector.Contains('\0'))
            throw new ArgumentException("The selector contains a forbidden sequence.");
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);
}

public enum MobileOperation { Set, Add, Remove }

public sealed record MobileOfficeCommand(
    MobileOperation Operation,
    string Path,
    string? Type = null,
    Dictionary<string, string>? Properties = null,
    int? Index = null,
    string? Before = null,
    string? After = null);

public sealed record MobileCommandResult(bool Success, string? Message, IReadOnlyList<string> UnsupportedProperties);
