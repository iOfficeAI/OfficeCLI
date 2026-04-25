// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;

namespace OfficeCli;

static partial class CommandBuilder
{
    // Stub Commands for the early-dispatch trio (mcp/skills/install).
    // These never execute their SetAction during normal use — Program.cs
    // intercepts those args before System.CommandLine sees them. The stubs
    // exist purely so:
    //   1. `officecli --help` lists them in its Commands section (no longer
    //      missing 3 commands relative to `officecli help`).
    //   2. `officecli <cmd> --help` reaches SCL (Program.cs falls through
    //      on --help/-h) and prints the usage from EarlyDispatchHelp.
    // Keep the usage strings in EarlyDispatchHelp (CommandBuilder.Help.cs)
    // as the single source of truth; this file just re-emits them.
    // Short blurbs shown both in `officecli --help`'s Commands list and at
    // the top of `officecli <cmd> --help`. Detailed multi-line usage lives
    // in EarlyDispatchHelp and is surfaced via `officecli help <cmd>` (the
    // single source of truth for verbose usage). Each blurb ends with a
    // hint pointing there, so `<cmd> --help` users discover it.
    private static readonly Dictionary<string, string> StubBlurbs =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["mcp"]     = "Start the MCP stdio server, or register/unregister officecli with an MCP client. Run 'officecli help mcp' for full usage.",
            ["skills"]  = "Install agent skill definitions (Claude Code, Cursor, Copilot, ...). Run 'officecli help skills' for full usage.",
            ["install"] = "One-step setup: install binary + skills + MCP for detected agents. Run 'officecli help install' for full usage.",
        };

    internal static IEnumerable<Command> BuildIntegrationStubCommands()
    {
        foreach (var (name, blurb) in StubBlurbs)
        {
            var cmd = new Command(name, blurb);
            // SetAction only fires when the user invokes the stub WITHOUT
            // --help/-h (Program.cs short-circuits the normal flow, so this
            // path is rarely hit). When it does fire, print the verbose
            // usage so the user isn't left with a bare blurb.
            cmd.SetAction(_ =>
            {
                if (EarlyDispatchHelp.TryGetValue(name, out var lines))
                    foreach (var line in lines) Console.WriteLine(line);
                return 0;
            });
            yield return cmd;
        }
    }
}
