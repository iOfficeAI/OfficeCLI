# OfficeCli Plugin Protocol

**Status**: v1 — final draft. No backward-compatibility goal; all plugins are
pre-release and re-align with this document.
**Audience**: Plugin authors and OfficeCli contributors.

## 1. Motivation

OfficeCli's main repo focuses on three universal Office formats (`.docx`, `.xlsx`,
`.pptx`). To extend format support without bloating the main binary or coupling
external implementations to the main repo's license, format support is delivered
through **plugins** — independent sidecar processes discovered and invoked by the
main binary.

Concrete drivers:

- Legacy formats (`.doc`, `.rtf`, `.odt`) where some users need migration but the
  parser is heavy and the format is fading
- Regional formats (`.hwpx`, `.hwp`) maintained by communities outside the main team
- Export targets (`.pdf`, `.epub`) where the renderer library has size, license,
  or platform constraints that make in-tree bundling undesirable
- Proprietary implementations that need to stay out of the Apache-licensed main
  repo

## 2. Plugin Kinds

A plugin declares its **kind** in its manifest. Each kind has a fixed
responsibility, lifecycle, and IPC pattern. v1 defines three kinds.

### 2.1 `dump-reader` — read a foreign format, emit officecli commands

Used to **migrate** a foreign format into one of main's native formats
(`.docx`/`.xlsx`/`.pptx`). The output format is declared by the plugin's
manifest `target` field.

| Aspect | Value |
|---|---|
| Lifecycle | Short-lived (one shot) |
| Source file handle | Plugin (read-only) |
| Target file handle | Main (replays plugin's batch into a sibling native file) |
| Vocabulary | **Main's `<target>` command vocabulary** (no plugin-defined extensions) |
| IPC | None — plugin writes JSONL (one `BatchItem` per line) to stdout and exits |
| Output extension | Sibling `<source-stem>.<target>` next to the source |

Flow:

1. User invokes a command that opens a `.doc` file
2. Main checks for a sibling `<source-stem>.<target>` next to the source. If
   it exists and is newer than the source, main opens it directly and skips
   steps 3–5
3. Main spawns the plugin: `<plugin> dump <source>`
4. Plugin parses the source and **streams** `add`/`set`/`batch` items to stdout
   as JSONL (one JSON object per line, terminated by `\n`), then exits 0
5. Main creates a blank `<target>` skeleton, replays the batch line-by-line,
   and moves it to the sibling path. Subsequent invocations reuse the sibling

Edits target the sibling native file, not the original source. Source-side changes
invalidate the cache automatically via mtime comparison; delete the sibling to
force reconversion.

**Streaming requirement**: dump-reader plugins MUST emit one batch item per
line, flushed individually. Top-level JSON arrays (`[{...},{...}]`) are
rejected by main with `corrupt_batch`. Streaming gives the host's idle
watchdog (§5.6) per-item activity signal and bounds main's memory usage on
large source files.

### 2.2 `exporter` — convert native format to a foreign target

Used to **render** native content (`.docx`/`.xlsx`/`.pptx`) into a foreign output
file (e.g. `.pdf`). Single-direction, no editing.

| Aspect | Value |
|---|---|
| Lifecycle | Short-lived |
| Source file handle | Plugin (reads native file, read-only) |
| Target file handle | Plugin (writes foreign file) |
| Vocabulary | None — no commands exchanged |
| IPC | None — plain CLI invocation, diagnostics on stderr |

Flow:

1. User invokes a view mode that targets a foreign format (e.g.
   `officecli view <file> pdf --out <path>`). The mode name maps to the
   target extension.
2. Main resolves the `(from, to)` pair to a plugin
3. Main spawns the plugin with the source path and target path
4. Plugin reads the source (using its own libraries), writes the target
5. Plugin exits 0 if the target was written successfully

**Source path is read-only.** Exporters MUST NOT write to or modify the source
file. This is a hard requirement: main passes the source path directly without
snapshotting. Plugins that need a writable working copy MUST create their own
temp copy.

### 2.3 `format-handler` — own a foreign format end-to-end

Used to support a **first-class non-native format** (e.g. `.hwpx`, `.hwp`). The
plugin holds the file open for the entire session and handles all document
operations.

| Aspect | Value |
|---|---|
| Lifecycle | Long-lived (session duration) |
| Source file handle | Plugin (read-write, same file as target) |
| Target file handle | Same as source |
| Vocabulary | **Plugin-defined** (declared in manifest, snapshotted at session start) |
| IPC | stdin/stdout (long-lived); stderr for diagnostics + heartbeat |

Flow:

1. User invokes a command on a `.hwpx` file
2. Main resolves `.hwpx` to a `format-handler` plugin
3. Main spawns the plugin with the file path; main writes requests to the plugin's stdin and reads replies from its stdout
4. Plugin opens the file and serves JSONL frames on stdin/stdout
5. Main and plugin exchange the **open handshake** (§5.3) — plugin replies
   with its runtime capabilities and vocabulary snapshot
6. Main wraps the plugin in a `FormatHandlerProxy : IDocumentHandler`; every
   operation becomes an IPC message
7. On session end, main sends `close`; plugin flushes pending writes (if any)
   and exits

### 2.4 Reserved kinds

The following kinds are reserved for future use. Plugins MUST NOT declare them
in v1:

- `engine` — pluggable backend for an in-tree subsystem (e.g. PDF rendering,
  field refresh)
- `transformer` — converts one native format to another (e.g. `.docx → .pptx`)

A plugin MAY declare multiple kinds in a single binary (e.g. an exporter that is
also a dump-reader). See §4.

## 3. Plugin Discovery

When main needs a plugin for `(kind, ext)`, it searches in this fixed order. The
first match wins.

1. **Environment variable**: `$OFFICECLI_PLUGIN_<KIND>_<EXT>` (absolute path to
   the plugin executable). Example: `$OFFICECLI_PLUGIN_DUMP_READER_DOC`.
2. **User plugins directory**:
   `~/.officecli/plugins/<kind>/<ext>/plugin(.exe)`
3. **Bundled plugins directory** (next to the main executable):
   `<dir>/plugins/<kind>/<ext>/plugin(.exe)`
4. **PATH lookup**: an executable named `officecli-<kind>-<ext>` or
   `officecli-<ext>` (in that priority).

Path conventions:

- `<kind>` uses kebab-case (`dump-reader`, `format-handler`, `exporter`)
- `<ext>` is the file extension without the leading dot (`doc`, `hwpx`, `pdf`)
- On Windows, `(.exe)` is appended automatically when searching
- Symlinks are followed

Main caches discovery results per process invocation. Adding a plugin between
invocations is picked up immediately.

## 4. Manifest

Every plugin MUST respond to `<plugin> --info` by printing a single JSON object
to stdout and exiting 0. The object describes the plugin to the main binary.

### 4.1 Required fields

| Field | Type | Description |
|---|---|---|
| `name` | string | Stable identifier, kebab-case (e.g. `officecli-doc`) |
| `version` | string | SemVer of the plugin (e.g. `1.0.0`) |
| `protocol` | integer | Protocol major version this plugin implements. v1 plugins MUST set `1`. Main rejects mismatches with exit code 5. |
| `kinds` | array | One or more declared kinds (see §2). Common case: `["dump-reader"]` |
| `extensions` | array | File extensions this plugin handles, leading dot (`[".doc"]`) |
| `idle_timeout_seconds` | object | Idle-timeout budget per verb. See §4.2. |
| `runtime` | string | Declarative runtime tag for diagnostics only: `dotnet` / `native` / `go` / `rust` / `python` / `other`. Host does not branch on this. |

The `target` field is **required** for `dump-reader` and MUST be one of
`"docx"`, `"xlsx"`, `"pptx"`. The `vocabulary` field is **required** for
`format-handler` (§4.4).

### 4.2 `idle_timeout_seconds`

Idle-timeout budgets in seconds. Main's watchdog kills the plugin when no
activity (stdout byte / RPC reply / stderr heartbeat) is observed within
this many seconds. **Total wall-clock time is not bounded** — long-running
work is fine as long as the plugin keeps producing output.

```json
"idle_timeout_seconds": {
  "default": 60,
  "verbs": {
    "dump": 30,
    "export": 120,
    "save": 30
  }
}
```

Rules:

- `default` is mandatory (positive integer)
- `verbs` is optional; entries override `default` for that verb
- `0` is **not allowed in the manifest** (avoids silent never-kill). Users
  can opt out at runtime via the `OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS`
  environment variable (see below)
- Recommended defaults (informative, not normative):
  - `dump-reader.dump` — 30s (streaming emit keeps idle low)
  - `exporter.export` — 60s (long jobs should heartbeat; see §5.6)
  - `format-handler` per-verb — 30s for reads, 60s for mutations/save

**User override**: set `OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS=<n>` in the
host environment to override the manifest budget for every verb in that
invocation (`0` disables the watchdog entirely). The override is for the
human user debugging a hung plugin — plugins themselves do not see this
variable, and it does not propagate into the plugin subprocess.

### 4.3 Optional fields

| Field | Type | Description |
|---|---|---|
| `description` | string | Short human-readable description |
| `target` | string | Native format the plugin produces (`"docx"`/`"xlsx"`/`"pptx"`). Required for `dump-reader`. |
| `tier` | string | Free-form tier identifier (`basic`/`pro`/`enterprise`) |
| `supports` | array | Capability tags (e.g. `["tables","images","fields"]`) |
| `limits` | object | Plugin-imposed limits (e.g. `{"maxFileSizeMb": 200}`) |
| `homepage` | string | URL |
| `license` | string | SPDX identifier |

### 4.4 Vocabulary (format-handler only)

Format-handler plugins MUST declare the vocabulary their proxied document model
exposes:

```json
"vocabulary": {
  "addable_types": ["page", "annotation", "formfield", "outline-item"],
  "settable_props": {
    "annotation": ["type", "rect", "color", "contents", "author", "opacity"],
    "page": ["rotation", "mediaBox"],
    "formfield": ["value", "readOnly"]
  },
  "path_segments": ["/page[N]", "/page[N]/annotation[M]", "/formfield[<name>]"]
}
```

Manifest vocabulary is used for **discovery and help output**. At session
start, the plugin returns a runtime **vocabulary snapshot** in the open
handshake reply (§5.3), which may differ from the manifest (e.g. extra
aliases). The host trusts the snapshot for validation.

**Vocabulary is documentation, not a runtime gate**: main does not reject
commands that fall outside the declared vocabulary. Plugins self-report
unsupported keys via the `set` reply's `unsupported_properties` list. This
follows the project-wide "handler-as-truth" principle.

### 4.5 Example manifests

`officecli-doc` (dump-reader):
```json
{
  "name": "officecli-doc",
  "version": "1.0.0",
  "protocol": 1,
  "kinds": ["dump-reader"],
  "extensions": [".doc"],
  "target": "docx",
  "runtime": "dotnet",
  "idle_timeout_seconds": { "default": 60, "verbs": { "dump": 30 } },
  "tier": "basic",
  "supports": ["paragraphs", "runs", "tables", "images", "lists"]
}
```

`officecli-pdf` (exporter):
```json
{
  "name": "officecli-pdf",
  "version": "0.1.0",
  "protocol": 1,
  "kinds": ["exporter"],
  "extensions": [".pdf"],
  "runtime": "dotnet",
  "idle_timeout_seconds": { "default": 60, "verbs": { "export": 120 } },
  "supports": ["from:docx", "from:xlsx", "from:pptx"]
}
```

`officecli-hwpx` (format-handler):
```json
{
  "name": "officecli-hwpx",
  "version": "0.9.0",
  "protocol": 1,
  "kinds": ["format-handler"],
  "extensions": [".hwpx"],
  "runtime": "dotnet",
  "idle_timeout_seconds": { "default": 30, "verbs": { "save": 60 } },
  "vocabulary": {
    "addable_types": ["paragraph", "run", "table", "image", "footnote"],
    "settable_props": { },
    "path_segments": [ ]
  }
}
```

## 5. Invocation

Beyond `--info`, each kind has its own subcommand surface.

### 5.1 dump-reader

```
<plugin> dump <source-file> [--media-dir <dir>]
```

- `<source-file>`: absolute path to the file to read
- `--media-dir`: optional scratch directory the plugin may use for transient
  files (e.g. extracted images referenced by command paths)

Main sets the `OFFICECLI_BIN` environment variable to the path of the running
officecli binary, so plugins that produce an intermediate `.docx` (e.g. via an
external converter) can shell out to `officecli dump <converted.docx>` and pipe
its output to stdout. Plugins that don't need this can ignore the variable.

**Output format**: JSONL — one JSON object per line, terminated by `\n`,
each line `flush`ed individually. Schema per line matches one entry of
`officecli batch --commands`:

```jsonl
{"command":"add","parent":"/body","type":"paragraph","props":{"text":"Hello"}}
{"command":"set","path":"/body/paragraph[1]","props":{"bold":"true"}}
```

A top-level JSON array on a single line is **rejected** with `corrupt_batch`.

Diagnostics go to stderr or `--log-file`. The plugin exits 0 on success; non-zero
codes follow §6.5.

### 5.2 exporter

```
<plugin> export <source-file> --out <target-file> [--options <json>]
```

- `<source-file>`: native format file (`.docx`/`.xlsx`/`.pptx`) — **read-only**
- `--out`: target path for the exported file
- `--options`: optional backend-specific options as a JSON string

The plugin MUST NOT write to or modify `<source-file>`. Main relies on this
to skip defensive snapshotting.

### 5.3 format-handler

```
<plugin> open <file>
```

The plugin reads request frames from **stdin** and writes reply frames to
**stdout** (one JSON object per line, terminated by `\n`). Diagnostic
output and heartbeat lines (§5.6) go on **stderr**. Anything the plugin
writes to stdout that is not a valid envelope is a plugin bug: main reports
it as `protocol_mismatch` and the session enters the broken state.

**Open handshake** (mandatory first exchange before any user command):

Main sends:
```json
{"protocol":1,"msg_type":"open","path":"<file>","editable":true}
```

Plugin replies:
```json
{"protocol":1,"msg_type":"ok","result":{
  "capabilities":{
    "commands":["add","set","get","query","remove","move","save","raw","raw-set"],
    "features":["save","extract-binary"]
  },
  "vocabulary":{
    "addable_types":[...],
    "settable_props":{...},
    "path_segments":[...]
  }
}}
```

Failure to handshake within the verb's idle timeout terminates the session.
The host caches the returned capabilities and vocabulary; subsequent
commands not present in `commands` are short-circuited with
`unsupported_command` without round-tripping.

After handshake, each request gets exactly one reply before the next request
is sent (§6.2).

#### Proxied verbs

Request envelope (main → plugin):
```json
{"protocol":1,"msg_type":"command","command":"<verb>","args":{...},"props":{...}}
```

**Read path:**

| `command` | `args` keys | `result` shape on `ok` |
|---|---|---|
| `view` | `mode` (`text`/`annotated`/`outline`/`stats`/`issues`), `start`/`end`/`max_lines`/`cols`/`type`/`limit`/`format` | string (or JSON object when `format=json`); for `mode=issues`, an array of issue objects |
| `get` | `path`, `depth` | DocumentNode JSON object |
| `query` | `selector` | array of DocumentNode |
| `validate` | (none) | array of `{error_type,description,path,part}` |

**Mutation path** (envelope carries `args` and `props` separately; `props` is
the user's `--prop key=value` dictionary, always string-to-string):

| `command` | `args` keys | `props` | `result` shape on `ok` |
|---|---|---|---|
| `set` | `path` | yes | object `{"unsupported_properties":["key1",...]}` (empty array = all applied) |
| `add` | `parent_path`, `type`, optional `position` | yes | object `{"path":"...","unsupported_properties":[...]}` |
| `remove` | `path` | no | string or null — optional warning text (e.g. cells shifted) |
| `move` | `source_path`, optional `target_parent_path`, optional `position` | no | string — new path |
| `copy` | `source_path`, `target_parent_path`, optional `position` | no | string — new path |
| `raw` | `part_path`, optional `start_row`/`end_row`/`cols` | no | string — raw XML (or CSV-of-rows for spreadsheet parts) |
| `raw_set` | `part_path`, `xpath`, `action`, optional `xml` | no | null |
| `add_part` | `parent_part_path`, `part_type` | optional | object `{"rel_id":"...","part_path":"..."}` |
| `extract_binary` | `path`, `dest_path` | no | object `{"found":true,"content_type":"...","byte_count":N}` or `{"found":false}` |

`position` (when present) is `{"index":N}` OR `{"after":"<path>"}` OR
`{"before":"<path>"}` — at most one field set; all-null means append.

**Numeric tolerance**: `byte_count` and similar integer fields MUST be JSON
numbers with no fractional part. Hosts SHOULD accept either int or
double-encoded integer forms (`42` and `42.0`) to absorb runtime drift across
languages.

#### `save`

```json
{"protocol":1,"msg_type":"save"}
```

`save` is **normative for format-handler plugins that accept mutations**.
The plugin MUST flush all pending writes to disk before replying `ok`. A
no-op acknowledgement is non-conformant and breaks main's crash-recovery
expectations. `plugins lint` verifies that a mutation followed by `save` is
durable by reopening the file from disk after the reply.

#### `close`

```json
{"protocol":1,"msg_type":"close"}
```

Plugin acknowledges with `ok`, flushes (implicit `save` if mutations were
applied without an explicit `save`), and exits 0.

### 5.4 Universal options

Each plugin subcommand SHOULD accept:

- `--log-file <path>`: append diagnostic output here instead of stderr
- `--quiet`: suppress non-error output

These are plugin-side conventions. The host's own idle-watchdog override is
the `OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS` env var (§4.2) — host does not
forward CLI flags into the plugin process for timeout purposes.

### 5.5 Cross-runtime conventions

To keep .NET / Go / Rust / native plugins interchangeable, all plugins MUST:

- Emit UTF-8 **without** BOM on stdout and stderr
- Use `\n` (not `\r\n`) as line separator on all platforms, including Windows
- Use **snake_case** for all JSON keys (manifest, IPC envelopes, error bodies)
- Return one of the documented exit codes (§6.5); non-zero codes that are
  not documented are reported as `internal_error`

### 5.6 Idle-timeout watchdog & heartbeat

Main runs a watchdog thread for every spawned plugin process:

- Any byte written to stdout (dump-reader, format-handler reply) **resets** the
  idle timer
- A line on stderr matching `{"heartbeat":true}` (optionally with extra
  fields) **resets** the idle timer without producing diagnostic noise. The
  heartbeat line is consumed by the watchdog and not surfaced to the user
- When `now - last_activity > idle_timeout`, main `Kill(entire_process_tree)`
  and reports `plugin_idle_timeout` (exit code 6)
- `--timeout 0` disables the watchdog; manifest cannot disable

Long opaque operations (exporter rendering, format-handler `save` on large
files) SHOULD emit periodic heartbeats. Plugins that stream output
naturally (dump-reader JSONL) do not need heartbeats.

## 6. IPC Protocol

Only `format-handler` exchanges live messages with main; the framing below
applies to that kind. (`dump-reader` and `exporter` are short-lived and use the
simpler stdout / exit-code contracts described in §5.1 and §5.2.)

### 6.1 Transport

Three standard streams, no auxiliary IPC channel:

- **stdin** — main writes request envelopes here, plugin reads them
- **stdout** — plugin writes reply envelopes here, main reads them
- **stderr** — plugin writes diagnostics and heartbeat lines here (§5.6)

The choice is deliberate: stdin/stdout is the same shape `dump-reader` and
`exporter` already use, every language has it built-in (no `NamedPipeClient`
or `UnixStream` wrapper to learn), and it sidesteps macOS's 104-byte
socket-path limit. The trade-off is one rule plugins MUST follow: stdout
carries protocol frames only — debug output goes to stderr or
`--log-file`. Main does not defend against polluted stdout; non-envelope
content is reported as `protocol_mismatch` and the session enters broken.

### 6.2 Framing & concurrency

UTF-8 text without BOM. One JSON object per line, terminated by `\n`. The
protocol is **request/response**: every client message receives exactly one
server reply before the next message is sent. For `format-handler`, **main
is the client** and **plugin is the server**.

Main MUST serialize requests per session. Callers in main that share a
single `FormatHandlerSession` MUST go through the session's internal mutex;
plugins MAY assume one request is in flight at a time.

### 6.3 Message envelope

Every message MUST include:

```json
{
  "protocol": 1,
  "msg_type": "<type>",
  ... type-specific fields ...
}
```

### 6.4 Message types

#### Request types (client → server)

| `msg_type` | Body |
|---|---|
| `open` | `{ "path": "<file>", "editable": <bool> }` (handshake, §5.3) |
| `command` | `{ "command": "add"\|"set"\|..., "args": {...}, "props": {...} }` |
| `save` | `{}` (normative flush, §5.3) |
| `close` | `{}` |
| `ping` | `{}` (liveness check; resets idle timer) |

#### Response types (server → client)

| `msg_type` | Body |
|---|---|
| `ok` | `{ "result": <value-or-null> }` |
| `error` | `{ "error": { "code": "<code>", "message": "...", "detail": "..." } }` |

#### Server-pushed events (format-handler only)

| `msg_type` | Body |
|---|---|
| `event` | `{ "kind": "warning"\|"info", "message": "..." }` |

Events are unsolicited and do not consume a reply slot; main MAY ignore them.

### 6.5 Exit codes

When a plugin process terminates:

| Code | Meaning |
|---|---|
| `0` | Success |
| `2` | Corrupt input file |
| `3` | Feature unsupported in this build |
| `4` | License expired |
| `5` | Protocol mismatch |
| `6` | Idle timeout (host-imposed; plugins do not emit this themselves) |
| `64`-`78` | Reserved (sysexits.h) |
| other | Plugin bug; main reports as `internal_error` |

### 6.6 Error codes (in `error.code`)

Plugins SHOULD use these codes when applicable:

| Code | Meaning |
|---|---|
| `invalid_request` | Malformed message |
| `unsupported_command` | Recognized message but unimplemented |
| `unsupported_feature` | Recognized command but feature not in this build |
| `invalid_argument` | Argument failed validation |
| `not_found` | Target path/element does not exist |
| `corrupt_input` | Source file is malformed or unreadable |
| `corrupt_batch` | dump-reader output is not valid JSONL |
| `license_expired` | Commercial plugin's license check failed |
| `protocol_mismatch` | Manifest protocol version differs from main's |
| `plugin_idle_timeout` | Host watchdog fired |
| `plugin_stream_closed` | stdin/stdout reached EOF before handshake or mid-session |
| `internal_error` | Catch-all for plugin bugs |

Codes are extensible; main treats unknown codes as `internal_error`.

### 6.7 Session lifecycle state machine

```
                       spawn process
   (none) ─────────────────────────────────► spawning
                                              │
                              open handshake  │ idle timer running
                              succeeded on    │
                              stdin/stdout    ▼
                                            ready
                                              │
                       command request sent   │ command reply received
                              ────────────►   │ ◄────────────
                                              ▼
                                             busy
                                              │
                                              │ stdin write failure
                                              │ OR stdout EOF / read failure
                                              │ OR idle timeout
                                              │ OR malformed reply
                                              ▼
                                            broken
                                              │
                                              │ Dispose / Kill
                                              ▼
                                            closed
```

Rules:

- Any IO failure or watchdog kill transitions to **broken**. Once broken,
  subsequent `Send` calls fail fast with `plugin_stream_closed`; the
  session is not auto-respawned (callers Dispose and re-open if needed)
- `close` reply transitions cleanly to **closed**
- The process is `Kill(entire_process_tree)`'d on transition to **closed**
  if it has not exited within 2 seconds of `close` reply (or immediately on
  transition from **broken**)

## 7. Vocabulary Contract

### 7.1 Universal protocol shell (all kinds)

These elements are stable across all plugins and all kinds:

- Message envelope shape (§6.3)
- Command verbs: `add`, `set`, `remove`, `move`, `get`, `query`, `batch`,
  `raw_set`
- Path syntax: `/segment[N]` with `[N]` 1-based index OR `[<name>]` named
  reference
- Error and exit code namespaces (extensible)

### 7.2 Per-format vocabulary

The specific **types** (`paragraph`/`page`/`cell`/...), **property names**
(`bold`/`fontsize`/`rect`/...), and **value formats** (`12pt`/`#FF0000`/...) are
not universal. They depend on which document model is at the other end:

- For `dump-reader`, the receiving model is main's `WordprocessingDocument` (or
  the spreadsheet/presentation equivalent for non-docx targets), so the
  vocabulary is main's `<target>` vocabulary (published as
  `schemas/word-vocabulary.json` etc.)
- For `format-handler`, the model is the plugin's own; the plugin declares its
  vocabulary in the manifest and reaffirms it via the open handshake
- For `exporter`, there is no command vocabulary

## 8. Installation

The protocol does **not** mandate any installation mechanism. As long as the
plugin executable ends up at one of the discovery paths (§3), it works.

Common installation channels:

- **Manual**: download a release archive, extract to `~/.officecli/plugins/...`
- **Bundled distribution**: main's release archive includes a `plugins/`
  directory next to the executable
- **Built-in installer** (recommended for users): `officecli plugins install <name>`
- **Package managers**: `dotnet tool install`, `winget`, `brew`, `apt`, `scoop`
- **Enterprise deployment**: place binaries via IT distribution

The built-in installer consults a registry (default:
`https://officecli.ai/plugins/registry.json`; configurable for private mirrors)
which lists approved plugins, versions, download URLs, and SHA-256 hashes.

## 9. Writing a Plugin

### 9.1 Minimum dump-reader (C#)

```csharp
using System.Text.Json;

if (args[0] == "--info") {
    Console.WriteLine(JsonSerializer.Serialize(new {
        name = "officecli-doc-minimal",
        version = "0.0.1",
        protocol = 1,
        kinds = new[] { "dump-reader" },
        extensions = new[] { ".doc" },
        target = "docx",
        runtime = "dotnet",
        idle_timeout_seconds = new { @default = 30 }
    }));
    return 0;
}

// args: dump <source-file>
string sourcePath = args[1];

// Parse source file (your library here) and emit one JSON object per line.
// Flush each line individually so main's idle watchdog sees activity.
var stdout = Console.Out;
stdout.WriteLine(JsonSerializer.Serialize(new {
    command = "add",
    parent = "/body",
    type = "paragraph",
    props = new { text = "Hello from .doc" }
}));
stdout.Flush();
// ... more items ...
return 0;
```

### 9.2 Minimum exporter (Go)

```go
package main

import (
    "encoding/json"
    "fmt"
    "os"
    "os/exec"
)

func main() {
    if len(os.Args) > 1 && os.Args[1] == "--info" {
        json.NewEncoder(os.Stdout).Encode(map[string]any{
            "name":       "officecli-pdf-min",
            "version":    "0.0.1",
            "protocol":   1,
            "kinds":      []string{"exporter"},
            "extensions": []string{".pdf"},
            "runtime":    "go",
            "idle_timeout_seconds": map[string]any{
                "default": 60,
                "verbs":   map[string]int{"export": 120},
            },
        })
        return
    }

    // args: export <source-file> --out <target-file>
    // MUST NOT write to source-file.
    source := os.Args[2]
    var target string
    for i, a := range os.Args {
        if a == "--out" && i+1 < len(os.Args) {
            target = os.Args[i+1]
        }
    }

    // Heartbeat on stderr for long jobs:
    go func() {
        for {
            fmt.Fprintln(os.Stderr, `{"heartbeat":true}`)
            time.Sleep(20 * time.Second)
        }
    }()

    cmd := exec.Command("soffice", "--headless", "--convert-to", "pdf",
        "--outdir", "/tmp/officecli-pdf", source)
    if err := cmd.Run(); err != nil {
        fmt.Fprintln(os.Stderr, err)
        os.Exit(3)
    }
    // ... move output to target ...
}
```

### 9.3 Minimum format-handler (C#, sketch)

```csharp
// args: open <file>
// stdin = requests from main, stdout = replies to main,
// stderr = diagnostics + heartbeat.
var stdin  = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false))
{
    NewLine = "\n",
    AutoFlush = true,
};

while (true) {
    var line = stdin.ReadLine();
    if (line == null) break;
    var msg = JsonNode.Parse(line)!;
    switch ((string)msg["msg_type"]!) {
        case "open":
            // load file, return capabilities + vocabulary snapshot
            stdout.WriteLine(JsonSerializer.Serialize(new {
                protocol = 1,
                msg_type = "ok",
                result = new {
                    capabilities = new {
                        commands = new[] { "get", "set", "save" },
                        features = Array.Empty<string>()
                    },
                    vocabulary = /* ... */ new {}
                }
            }));
            break;
        case "save":
            // MUST actually flush to disk before replying ok
            File.WriteAllBytes(filePath, currentBytes);
            stdout.WriteLine("""{"protocol":1,"msg_type":"ok","result":null}""");
            break;
        case "close":
            stdout.WriteLine("""{"protocol":1,"msg_type":"ok","result":null}""");
            return 0;
        // ... command dispatch ...
    }
}
```

## 10. Stability Commitments

### 10.1 Main → Plugins

Once protocol v1 is ratified, main commits to:

1. **Protocol shell** is stable for v1. Adding new optional message types is
   allowed; removing or changing types requires a v2 bump.
2. **Native vocabulary** (relevant to `dump-reader`): additions allowed;
   deletions or renames require a deprecation cycle of at least two minor
   releases with the old name accepted as an alias.
3. **Path syntax** does not change.
4. **Error/exit code semantics** do not change. Adding new codes is allowed.
5. **Schema files** (`schemas/word-vocabulary.json`, etc.) are released
   alongside main and follow the same versioning.

### 10.2 Plugins → Main

Plugin authors should:

1. Treat `--info` output schema as stable per protocol major version.
2. Implement graceful degradation when main lacks expected capabilities.
3. Provide a meaningful exit code on failure (don't silently exit 1 for every
   error).
4. Avoid writing to paths other than `--media-dir`, the declared output file,
   or temp files the plugin owns.

## 11. FAQ

**Q: Can plugins be in any language?**
A: Yes. The protocol is JSONL over stdin/stdout. Any language with
subprocess and standard-stream support works. .NET plugins can optionally use the
`OfficeCli.Contracts` NuGet package for type-safe types.

**Q: How does main know which plugin to use when several are installed?**
A: Discovery order (§3) is fixed and first-match-wins. For multiple installed
plugins for the same extension, users select via env var or explicit
`--plugin` flag.

**Q: Can a plugin be closed-source / commercial?**
A: Yes. Plugins are independent binaries with their own license. License
check failures exit 4 (`license_expired`).

**Q: What if the plugin crashes?**
A: Main detects non-zero exit and surfaces a clear error. Partial state in
main's in-memory document is discarded; no corrupt files are written.

**Q: What if the plugin hangs?**
A: Main's idle watchdog (§5.6) kills it when no output is observed within
the manifest-declared `idle_timeout_seconds`. Long jobs heartbeat on stderr
to stay alive.

**Q: Why no total wall-clock timeout?**
A: Large .doc files legitimately take minutes to dump; Word-interop PDF
export of large workbooks can take hours. A wall-clock cap punishes correct
behavior. Idle timeout catches actual hangs without false positives.

**Q: How does this differ from MCP?**
A: MCP exposes officecli to AI clients; plugins extend officecli's format
support. The two are complementary.

## 12. Versioning

This document tracks **protocol** version, distinct from main repo version.

- v1.x: Additive changes only (new optional fields, new message types, new
  error codes). Backward-compatible.
- v2.x: Breaking changes (removed/renamed fields, changed semantics).

Main repo declares supported protocol version(s) via `officecli --version`.
Plugins declare their target protocol in manifest. Main rejects plugins
whose major protocol version differs from main's supported version, exiting
the plugin process with code 5 and surfacing `protocol_mismatch` to the user.

## 13. Open Questions (post-v1)

- Should `format-handler` plugins support concurrent multi-document sessions in
  one process? (v1: no, one process per open document)
- Should the registry support package signing? (Likely yes for v1.1)
- Should `capabilities` queries return JSON Schema fragments inline, or only
  list names? (Currently: names; consider inline schema in v1.1)
- Host-driven session pooling for format-handler (kill idle sessions to free
  memory). Not in v1; revisit if process count becomes a real problem.

---

*This document is the source of truth for the OfficeCli Plugin Protocol v1.
Pre-release plugins re-align with this document; post-ratification changes
follow §10 and §12.*
