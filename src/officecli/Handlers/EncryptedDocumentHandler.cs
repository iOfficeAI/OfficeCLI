// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

/// <summary>
/// Wraps a normal <see cref="IDocumentHandler"/> that is operating on a
/// <i>decrypted</i> copy of a password-protected OOXML file, and transparently
/// re-encrypts on save.
///
/// <para>The native handlers (Word/Excel/PowerPoint) read and save a document
/// in place through a backing <c>FileStream</c> and have no notion of
/// encryption. Rather than teach each of them MS-OFFCRYPTO, the factory decrypts
/// the package to a temporary plaintext file, opens an ordinary handler on that
/// temp file, and hands it to this decorator. Every operation flows straight
/// through to the inner handler; only the <i>lifecycle</i> is intercepted:
/// after the inner handler flushes plaintext to the temp file, this re-encrypts
/// it under the original password and writes the result back to the original
/// (encrypted) path. On dispose the temp plaintext is shredded-by-delete.</para>
///
/// <para>Read-only sessions (<c>editable: false</c>, e.g. <c>view</c>/<c>get</c>)
/// never write the original back — the temp file is just decrypted-for-reading
/// and removed on dispose.</para>
/// </summary>
internal sealed class EncryptedDocumentHandler : IDocumentHandler
{
    private readonly IDocumentHandler _inner;
    private readonly string _tempPlainPath;
    private readonly string _originalEncryptedPath;
    private readonly string _password;
    private readonly bool _editable;
    private bool _disposed;

    public EncryptedDocumentHandler(
        IDocumentHandler inner, string tempPlainPath, string originalEncryptedPath,
        string password, bool editable)
    {
        _inner = inner;
        _tempPlainPath = tempPlainPath;
        _originalEncryptedPath = originalEncryptedPath;
        _password = password;
        _editable = editable;
    }

    /// <summary>Re-encrypt the current plaintext temp file back to the original
    /// path. No-op for read-only sessions.</summary>
    private void ReEncryptToOriginal()
    {
        if (!_editable) return;
        byte[] plain = File.ReadAllBytes(_tempPlainPath);
        byte[] encrypted = MsOffCrypto.Encrypt(plain, _password);
        File.WriteAllBytes(_originalEncryptedPath, encrypted);
    }

    // === Lifecycle (intercepted) ===
    public void Save()
    {
        _inner.Save();              // inner flushes plaintext to the temp file
        ReEncryptToOriginal();      // then we re-seal it back to the original
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            _inner.Dispose();       // inner's final save lands in the temp file
            ReEncryptToOriginal();
        }
        finally
        {
            try { File.Delete(_tempPlainPath); } catch { /* best-effort cleanup */ }
        }
    }

    // === Everything else: straight delegation to the inner handler ===
    public string ViewAsText(int? s = null, int? e = null, int? m = null, HashSet<string>? c = null) => _inner.ViewAsText(s, e, m, c);
    public string ViewAsAnnotated(int? s = null, int? e = null, int? m = null, HashSet<string>? c = null) => _inner.ViewAsAnnotated(s, e, m, c);
    public string ViewAsOutline() => _inner.ViewAsOutline();
    public string ViewAsStats() => _inner.ViewAsStats();
    public JsonNode ViewAsStatsJson() => _inner.ViewAsStatsJson();
    public JsonNode ViewAsOutlineJson() => _inner.ViewAsOutlineJson();
    public JsonNode ViewAsTextJson(int? s = null, int? e = null, int? m = null, HashSet<string>? c = null) => _inner.ViewAsTextJson(s, e, m, c);
    public List<DocumentIssue> ViewAsIssues(string? t = null, int? l = null) => _inner.ViewAsIssues(t, l);
    public DocumentNode Get(string path, int depth = 1) => _inner.Get(path, depth);
    public List<DocumentNode> Query(string selector) => _inner.Query(selector);
    public List<string> Set(string path, Dictionary<string, string> properties) => _inner.Set(path, properties);
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties) => _inner.Add(parentPath, type, position, properties);
    public string? Remove(string path, Dictionary<string, string>? properties = null) => _inner.Remove(path, properties);
    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position, Dictionary<string, string>? properties = null) => _inner.Move(sourcePath, targetParentPath, position, properties);
    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position) => _inner.CopyFrom(sourcePath, targetParentPath, position);
    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null) => _inner.Raw(partPath, startRow, endRow, cols);
    public void RawSet(string partPath, string xpath, string action, string? xml) => _inner.RawSet(partPath, xpath, action, xml);
    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null) => _inner.AddPart(parentPartPath, partType, properties);
    public List<ValidationError> Validate() => _inner.Validate();
    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount) => _inner.TryExtractBinary(path, destPath, out contentType, out byteCount);
}
