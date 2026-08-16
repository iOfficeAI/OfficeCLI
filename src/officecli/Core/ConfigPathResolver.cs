// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

internal static class ConfigPathResolver
{
    internal static string ResolveCurrent()
    {
        var userHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var legacyDirectory = Path.Combine(userHome, ".officecli");
        if (!OperatingSystem.IsLinux())
            return legacyDirectory;

        var xdgConfigHome = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME");
        var xdgDirectory = XdgDirectory(userHome, xdgConfigHome);
        return Resolve(
            userHome,
            xdgConfigHome,
            isLinux: true,
            Directory.Exists(legacyDirectory),
            File.Exists(Path.Combine(legacyDirectory, "config.json")),
            File.Exists(Path.Combine(xdgDirectory, "config.json")));
    }

    internal static string Resolve(
        string userHome,
        string? xdgConfigHome,
        bool isLinux,
        bool legacyDirectoryExists,
        bool legacyConfigExists,
        bool xdgConfigExists)
    {
        var legacyDirectory = Path.Combine(userHome, ".officecli");
        if (!isLinux)
            return legacyDirectory;

        var xdgDirectory = XdgDirectory(userHome, xdgConfigHome);
        if (legacyConfigExists)
            return legacyDirectory;

        // A fresh XDG install may still create ~/.officecli later for logs,
        // plugins, or caches. Once its config exists, do not switch paths.
        if (xdgConfigExists)
            return xdgDirectory;

        return legacyDirectoryExists ? legacyDirectory : xdgDirectory;
    }

    private static string XdgDirectory(string userHome, string? xdgConfigHome)
    {
        var configRoot = string.IsNullOrWhiteSpace(xdgConfigHome) ||
                         !Path.IsPathRooted(xdgConfigHome)
            ? Path.Combine(userHome, ".config")
            : xdgConfigHome;
        return Path.Combine(configRoot, "officecli");
    }
}
