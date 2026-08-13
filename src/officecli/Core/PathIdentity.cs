// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// One document ↔ one identity string. Every per-file IPC channel
/// (resident pipe/lock, watch pipe/marker) hashes the document path to name
/// itself, so two commands that name the SAME file must produce the SAME
/// string — otherwise they talk past each other: a second resident opens for
/// a file already held by one (duplicated in-memory edits), or a mutation's
/// watch notification goes to a pipe nobody listens on (preview never updates).
///
/// <see cref="System.IO.Path.GetFullPath(string)"/> alone is not enough: it
/// makes a path absolute but does NOT resolve symlinks. On macOS
/// <c>/tmp</c> is a link to <c>/private/tmp</c> and the process CWD is
/// reported already-resolved, so <c>watch test.pptx</c> (→ /private/tmp/…)
/// and <c>set /tmp/…/test.pptx</c> hash differently for the same file.
/// </summary>
internal static class PathIdentity
{
    // Bounds a symlink cycle (a → b → a) and pathological chains.
    private const int MaxLinkDepth = 40;

    /// <summary>
    /// Absolute, symlink-resolved path for <paramref name="filePath"/>. The
    /// file need not exist — unresolvable segments are kept verbatim. Never
    /// throws; falls back to the plain absolute path.
    /// </summary>
    public static string Canonical(string filePath)
    {
        try
        {
            var full = Path.GetFullPath(filePath);
            var dir = Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(dir)) return full;
            var resolved = Path.Combine(ResolveDirectory(dir, 0), Path.GetFileName(full));
            // The leaf itself may be a link (e.g. a symlinked report.docx).
            var leafInfo = new FileInfo(resolved);
            if (leafInfo.Exists)
            {
                var leafTarget = leafInfo.ResolveLinkTarget(returnFinalTarget: true);
                if (leafTarget != null) resolved = leafTarget.FullName;
            }
            return resolved;
        }
        catch
        {
            try { return Path.GetFullPath(filePath); }
            catch { return filePath; }
        }
    }

    private static string ResolveDirectory(string dir, int depth)
    {
        if (depth >= MaxLinkDepth) return dir;
        var info = new DirectoryInfo(dir);
        if (info.Exists)
        {
            var target = info.ResolveLinkTarget(returnFinalTarget: true);
            if (target != null) return ResolveDirectory(target.FullName, depth + 1);
        }
        var parent = Path.GetDirectoryName(dir);
        if (string.IsNullOrEmpty(parent) || parent == dir) return dir;
        var resolvedParent = ResolveDirectory(parent, depth + 1);
        return resolvedParent == parent
            ? dir
            : Path.Combine(resolvedParent, Path.GetFileName(dir));
    }
}
