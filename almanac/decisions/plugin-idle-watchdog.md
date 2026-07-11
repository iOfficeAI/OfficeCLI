---
title: "Plugin Idle Watchdog"
summary: "Plugin processes are timed out by lack of activity, not total runtime, with manifest budgets and a user override for debugging."
topics: [decisions, plugins, runtime]
sources:
  - id: plugin-protocol
    type: file
    path: plugins/plugin-protocol.md
  - id: plugin-process
    type: file
    path: src/officecli/Core/Plugins/PluginProcess.cs
  - id: plugin-manifest
    type: file
    path: src/officecli/Core/Plugins/PluginManifest.cs
---

OfficeCLI uses an activity-based watchdog for plugin subprocesses. A plugin may run for a long time, but it must keep producing protocol output, diagnostics, or heartbeat lines within its idle budget. This lets large conversions and exports continue while still killing hung plugins that stop communicating [@plugin-protocol] [@plugin-process].

## Context

Plugins are sidecar processes used by the [Plugin System](../architecture/plugins/plugin-system). Dump-readers and exporters are short-lived, while format-handlers can hold a long-lived editing session [@plugin-protocol].

Those processes can hang in external parsers, renderers, license checks, or file I/O. A wall-clock timeout would be too blunt for large documents, but no timeout would leave OfficeCLI waiting forever.

## Decision

The timeout measures idle time. The protocol says any stdout byte resets the timer, and a stderr line matching `{"heartbeat":true}` resets it without being surfaced as user diagnostics [@plugin-protocol]. The shared `PluginProcess` runner implements that rule by updating an activity timestamp from stdout lines, heartbeat stderr lines, and non-heartbeat stderr output [@plugin-process].

When the gap since last activity exceeds the budget, the runner kills the entire process tree and reports an idle timeout [@plugin-process]. The protocol maps that condition to `plugin_idle_timeout` with exit code 6 [@plugin-protocol].

Budgets come from the plugin manifest. `idle_timeout_seconds.default` is the fallback, and optional per-verb entries override it [@plugin-protocol] [@plugin-manifest]. The manifest model resolves a timeout for each verb and falls back to a safe default when the manifest is silent [@plugin-manifest].

Users can override the manifest with `OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS`. A non-negative value replaces the manifest budget for the invocation, and `0` disables the watchdog for debugging [@plugin-protocol] [@plugin-manifest].

## Status

The decision is implemented in `PluginProcess.Run`, `PluginManifest.ResolveIdleTimeout`, and the protocol document [@plugin-process] [@plugin-manifest] [@plugin-protocol].

## Consequences

Plugin authors must treat stdout and stderr as liveness channels. Dump-readers naturally reset the watchdog by streaming JSONL output, while exporters and long format-handler operations should emit heartbeat lines during opaque work [@plugin-protocol].

The host intentionally does not bound total wall-clock runtime. A large job can run for minutes if it keeps making observable progress [@plugin-process].

The main risk is polluted streams. Heartbeats belong on stderr, and protocol frames belong on stdout. The protocol warns that non-envelope stdout in format-handler sessions is a plugin bug and can break the session [@plugin-protocol].
