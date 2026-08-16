// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

var home = Path.Combine(Path.GetPathRoot(Environment.CurrentDirectory)!, "home", "officecli-test");
var xdg = Path.Combine(Path.GetPathRoot(Environment.CurrentDirectory)!, "xdg", "config");
var legacy = Path.Combine(home, ".officecli");

AssertPath(
    Path.Combine(xdg, "officecli"),
    ConfigPathResolver.Resolve(
        home, xdg, isLinux: true,
        legacyDirectoryExists: false,
        legacyConfigExists: false,
        xdgConfigExists: false),
    "fresh Linux install honors XDG_CONFIG_HOME");

AssertPath(
    Path.Combine(home, ".config", "officecli"),
    ConfigPathResolver.Resolve(
        home, null, isLinux: true,
        legacyDirectoryExists: false,
        legacyConfigExists: false,
        xdgConfigExists: false),
    "fresh Linux install defaults to ~/.config");

AssertPath(
    Path.Combine(home, ".config", "officecli"),
    ConfigPathResolver.Resolve(
        home, "   ", isLinux: true,
        legacyDirectoryExists: false,
        legacyConfigExists: false,
        xdgConfigExists: false),
    "blank XDG_CONFIG_HOME defaults to ~/.config");

AssertPath(
    Path.Combine(home, ".config", "officecli"),
    ConfigPathResolver.Resolve(
        home, Path.Combine("relative", "config"), isLinux: true,
        legacyDirectoryExists: false,
        legacyConfigExists: false,
        xdgConfigExists: false),
    "relative XDG_CONFIG_HOME defaults to ~/.config");

AssertPath(
    legacy,
    ConfigPathResolver.Resolve(
        home, xdg, isLinux: true,
        legacyDirectoryExists: true,
        legacyConfigExists: false,
        xdgConfigExists: false),
    "existing Linux ~/.officecli directory keeps the legacy location");

AssertPath(
    legacy,
    ConfigPathResolver.Resolve(
        home, xdg, isLinux: false,
        legacyDirectoryExists: true,
        legacyConfigExists: false,
        xdgConfigExists: true),
    "non-Linux platforms keep the legacy location");

AssertPath(
    Path.Combine(xdg, "officecli"),
    ConfigPathResolver.Resolve(
        home, xdg, isLinux: true,
        legacyDirectoryExists: true,
        legacyConfigExists: false,
        xdgConfigExists: true),
    "an XDG config stays selected after another feature creates ~/.officecli");

AssertPath(
    legacy,
    ConfigPathResolver.Resolve(
        home, xdg, isLinux: true,
        legacyDirectoryExists: true,
        legacyConfigExists: true,
        xdgConfigExists: true),
    "a legacy config wins when both config files exist");

Console.WriteLine("Config path resolver tests passed.");

static void AssertPath(string expected, string actual, string scenario)
{
    if (expected == actual) return;
    throw new InvalidOperationException(
        $"{scenario}: expected '{expected}', got '{actual}'");
}
