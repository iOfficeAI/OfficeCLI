// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Shared logic for invoking exporter plugins (docs/plugin-protocol.md §5.2):
/// resolution, subprocess invocation, exit-code mapping. Used by
/// `view &lt;file&gt; pdf` and any future caller that needs to convert a native
/// document to a foreign target via an installed exporter plugin.
///
/// Per §5.2 the plugin MUST NOT write to the source path, so main passes
/// the source path directly without snapshotting.
/// </summary>
public static class ExporterInvoker
{
    public sealed record ExportResult(string OutputPath, ResolvedPlugin Plugin, bool ResidentClosed);

    /// <summary>
    /// Resolve an exporter for (sourceExt, targetExt) and run it. On success,
    /// the target file exists at <paramref name="outPath"/> and the result
    /// reports which plugin handled it. On failure, throws CliException with
    /// an appropriate code (exporter_not_found, plugin_failed, ...).
    ///
    /// If a resident is holding the source file, it's closed first to release
    /// the exclusive lock; <see cref="ExportResult.ResidentClosed"/> indicates
    /// this happened so the caller can surface it to the user.
    /// </summary>
    public static ExportResult Run(string sourceFullPath, string targetExt, string outPath)
    {
        var sourceExt = Path.GetExtension(sourceFullPath).ToLowerInvariant();

        var plugin = Resolve(sourceExt, targetExt)
            ?? throw new CliException($"No exporter plugin found for {sourceExt} → {targetExt}.")
            {
                Code = "exporter_not_found",
                Suggestion = "Install an exporter plugin: `officecli plugins list` to see what's available, or see docs/plugin-protocol.md.",
            };

        bool residentClosed = false;
        if (ResidentClient.TryConnect(sourceFullPath, out _))
        {
            if (ResidentClient.SendCloseWithResponse(sourceFullPath, out _))
                residentClosed = true;
        }

        var idle = plugin.Manifest.ResolveIdleTimeout("export");
        var result = PluginProcess.Run(new PluginProcess.RunOptions
        {
            ExecutablePath = plugin.ExecutablePath,
            Arguments = new[] { "export", sourceFullPath, "--out", outPath },
            IdleTimeoutSeconds = idle,
        });

        if (result.IdleTimedOut)
            throw new CliException(
                $"Exporter plugin '{plugin.Manifest.Name}' produced no output for {idle}s — likely hung.")
            {
                Code = "plugin_idle_timeout",
                Suggestion = "Override with --timeout 0 or raise `idle_timeout_seconds.verbs.export` in the plugin's manifest. " +
                             "Long-running exporters should also emit `{\"heartbeat\":true}` on stderr periodically.",
            };

        if (result.ExitCode != 0)
            throw new CliException(
                $"Exporter plugin '{plugin.Manifest.Name}' failed (exit {result.ExitCode}): {Truncate(result.Stderr, 500)}")
            {
                Code = result.ExitCode switch
                {
                    2 => "corrupt_input",
                    3 => "unsupported_feature",
                    4 => "license_expired",
                    5 => "protocol_mismatch",
                    6 => "plugin_idle_timeout",
                    _ => "plugin_failed",
                },
            };

        if (!File.Exists(outPath))
            throw new CliException(
                $"Exporter plugin '{plugin.Manifest.Name}' reported success but no output file was written at {outPath}.")
            { Code = "plugin_contract_violation" };

        return new ExportResult(outPath, plugin, residentClosed);
    }

    /// <summary>
    /// Find an exporter for (source, target). Indexed by target extension (the
    /// plugin's declared extensions field); filtered by source via the manifest's
    /// supports list. A plugin missing supports is assumed to accept all native
    /// sources — conservative default for older manifests.
    /// </summary>
    public static ResolvedPlugin? Resolve(string sourceExt, string targetExt)
    {
        var p = PluginRegistry.FindFor(PluginKind.Exporter, targetExt);
        if (p is null) return null;

        if (p.Manifest.Supports is null || p.Manifest.Supports.Count == 0)
            return p;

        var sourceBare = sourceExt.TrimStart('.');
        if (p.Manifest.Supports.Any(s =>
                string.Equals(s, $"from:{sourceBare}", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, sourceExt, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, sourceBare, StringComparison.OrdinalIgnoreCase)))
            return p;

        return null;
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
