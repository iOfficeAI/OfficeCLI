# Threat model

## What this project does and where untrusted input enters
- OfficeCLI is a command-line tool (C#, .NET 10, single self-contained binary) that reads and writes `.docx`, `.xlsx` and `.pptx` files. It runs locally with the invoking user's privileges and is commonly driven by AI agents on behalf of a user. It is a single-user tool, meant for the user's own machine and account; it is not designed to be run on a host shared with untrusted local accounts.
- **The document file is untrusted.** Assume any `.docx`/`.xlsx`/`.pptx` passed to a command (`get`, `query`, `view`, `dump`, `validate`, `set`, `add`, `merge`, `import`, `batch`, …) was crafted by an attacker. This covers the zip container, every XML part, relationships (including external targets), embedded objects/images/fonts, and formulas evaluated by the built-in xlsx formula engine.
- **Rendered output is untrusted-derived.** `view html` / `view svg` and the `watch` live preview turn document content into HTML/SVG that is opened in a browser.
- **Remote resources.** Images and files given by URL go through `Core/ImageSource.cs` / `Core/FileSource.cs`, guarded by `Core/SsrfGuard.cs`.
- **Local servers.**
  - `watch` serves the preview over HTTP on a loopback TCP port (`Core/Watch/WatchServer.cs`), with Host/Origin checks against DNS rebinding and cross-site requests. Assume a malicious web page is open in the user's browser.
  - The resident server (`ResidentServer.cs`) accepts commands over a local named pipe.
  - Both are local-only conveniences for the invoking user. They are not a security boundary between different OS accounts on the same machine (see "Anything to leave alone").
- The command-line arguments and `--prop` values come from the user (or the user's agent) and are trusted, with one exception: they must not let content inside a document escalate into actions the user did not ask for.

## Components that matter most / least
- Most: document open/parse paths (`Handlers/DocumentHandlerFactory.cs`, `Handlers/Word|Excel|Pptx/`), zip handling, XML parsing, the formula evaluator, the HTML/SVG renderers, `SsrfGuard`, and the watch/resident servers.
- Less: write-only paths that only process the user's own arguments (`create`, `set` on values the user typed).
- Out of scope: `sdk/`, `npm/`, `skills/`, `examples/`, install scripts, and third-party plugins loaded through `officecli plugins` (users install them themselves and they run with the user's privileges by design).

## How to exercise it
- The image has the release binary at `/usr/local/bin/officecli`. `officecli help` lists commands; `officecli help <format> <element>` shows properties.
- Typical read paths on a crafted file: `officecli view evil.docx text`, `officecli view evil.docx html -o out.html`, `officecli get evil.xlsx /Sheet1/A1 --json`, `officecli query evil.pptx shape`, `officecli dump evil.docx`, `officecli validate evil.docx`.
- `examples/` contains about 100 valid sample documents to mutate as a starting point.
- Rebuild after a change: `dotnet build src/officecli/officecli.csproj -r linux-x64 --no-restore` (dependencies are restored in the image).
- Set `OFFICECLI_NO_AUTO_RESIDENT=1` (already set in the image) so each command opens the file itself instead of starting a background resident process.

## How you rate severity
- **Critical:** opening a crafted document leads to code execution, or to writing/overwriting files outside the paths the user named (e.g. zip-slip, path traversal through part names or relationships).
- **High:** opening a crafted document discloses local files or environment data (e.g. XXE, external relationship resolution); an `SsrfGuard` bypass that reaches internal or loopback addresses; a web page open in the user's browser can drive the `watch` server (read documents, run commands).
- **Medium:** script execution in `view html` / `watch` preview output (XSS) from document content; resource exhaustion (zip bombs, XML entity or recursion blowups, runaway formula evaluation) that hangs or crashes the process.
- **Low:** a plain crash, unhandled exception or wrong exit code on malformed input with no further impact.
- Please include a minimal crafted document (or a script that generates it) and the exact `officecli` command that triggers the issue. Patches should be minimal, follow the surrounding code style, and not add new dependencies.

## Anything to leave alone
- Commands that intentionally accept arbitrary content from the user: `raw-set`, `add-part`, `batch` scripts and `--prop` values the user supplied.
- Documents that OfficeCLI writes and that Microsoft Office then refuses to open, with no security impact (functional bugs, not vulnerabilities).
- Supply-chain or self-update concerns of `UpdateChecker` / release downloads, and anything that needs an attacker who can already run code as the same user.
- Anything that needs an attacker who already has another OS account on the same machine (connecting to the `watch` loopback port, or to / ahead of the resident and watch pipes). OfficeCLI is a single-user tool and does not defend against other local accounts; do not run `watch` or a resident on a host shared with accounts you do not trust.
