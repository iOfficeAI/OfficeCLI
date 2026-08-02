r"""
officecli — a thin Python shell over officecli's resident pipe.

It does ONE thing: forward a command to the running resident over its named
pipe and hand back the response. There is NO second vocabulary to learn: a
command is the same dict you'd put in an officecli `batch` list — e.g.
{"command":"set","path":"/Sheet1/A1","props":{"text":"Hello"}}. `send` forwards
one; `batch` forwards many in a single round-trip.

Two surfaces, by design:
  - bootstrap (infrequent): `create` / `open` spawn ONE CLI process — a file that
    isn't open yet (or doesn't exist yet) has no resident to talk to.
  - everything else (the hot path): `send` / `batch` are pure pipe round-trips,
    no per-command process spawn.

    import officecli
    with officecli.create("report.xlsx", "--force") as doc:   # make file + get handle
        doc.send({"command": "set", "path": "/Sheet1/A1",
                  "props": {"text": "Hello"}})
        print(doc.send({"command": "get", "path": "/Sheet1/A1"}))
        doc.send({"command": "save"})
    # ...or officecli.open("existing.xlsx") for a file that already exists.

The item keys are officecli's batch fields (command/op, path, parent, type,
index, after, before, to, selector, text, mode, depth, part, xpath, action,
xml) plus a nested `props` dict. Everything except command/op/props is
forwarded verbatim as a command argument; the resident dispatches it exactly
like the matching CLI command. See `officecli help` / the batch docs for the
field-and-prop reference — this shell adds none of its own.

Protocol (matches ResidentServer.cs / ResidentClient.cs):
  - pipe name : officecli-<SHA256(fullpath)[:16] uppercase>;
                fullpath upper-cased on macOS/Windows, left as-is on Linux.
  - unix path : $TMPDIR/CoreFxPipe_<name>  (+ "-ping");  $TMPDIR else /tmp
  - win path  : \\.\pipe\<name>            (+ "-ping")
  - framing   : one request line + one response line, UTF-8, '\n' terminated;
                one connection == one command.
  - request   : PascalCase {"Command","Args","Props","Json"}
  - response  : {"ExitCode","Stdout","Stderr"}
"""

import os
import sys
import json
import time
import socket
import hashlib
import shutil
import threading
import subprocess

# Mirror officecli's TryResident busy-delivery policy (CommandBuilder.cs): a
# generous connect timeout + a few retries with backoff, applied identically to
# every command. The reply read itself blocks (no timeout) — like officecli's
# PipeReadLine — trusting the resident to answer once our turn comes up in its
# serialized queue. Because retries only re-attempt the CONNECT (before the
# command executes), re-sending is safe even for mutations; there is no
# "read timed out, resend" path that could double-apply.
_BUSY_CONNECT_TIMEOUT = 30.0   # = ResidentBusyConnectTimeoutMs (30000)
_BUSY_MAX_RETRIES = 3          # = ResidentBusyMaxRetries
# = CommandBuilder's DefaultOpenIdleSeconds. `open` upgrades a reused resident
# (which `create` may have auto-started with a short 60s timeout) to the 12min
# interactive window, so a long editing session over an SDK handle isn't cut
# short by the create-time timeout.
_OPEN_IDLE_SECONDS = 12 * 60

_IS_WIN = sys.platform.startswith("win")
_IS_MAC = sys.platform == "darwin"
_builtin_open = open   # preserved; this module defines its own open() below

# officecli's official installer (README one-liner). install() shells out to it;
# the missing-CLI error points users at it / at install().
# Installer scripts: the d.officecli.ai mirror is primary; GitHub raw is only a
# fallback (same order as install.sh / install.ps1 themselves). The mirror is
# Cloudflare-fronted and reachable where raw.githubusercontent.com may be
# rate-limited or blocked.
_INSTALL_SH_MIRROR = "https://d.officecli.ai/install.sh"
_INSTALL_SH_GITHUB = "https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh"
_INSTALL_PS1_MIRROR = "https://d.officecli.ai/install.ps1"
_INSTALL_PS1_GITHUB = "https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1"
_MISSING_CLI = (
    "officecli CLI not found: {bin!r} is not on PATH nor in the default install "
    "location (~/.local/bin, or %LOCALAPPDATA%\\OfficeCLI on Windows). This SDK only forwards "
    "commands to the officecli binary, which must be installed separately. Install it:\n"
    "    python -m officecli install            # runs the official installer\n"
    "    # or: curl -fsSL " + _INSTALL_SH_MIRROR + " | bash\n"
    "Already installed elsewhere? pass binary=\"/path/to/officecli\"."
)


class OfficeCliError(Exception):
    """Raised on transport/process failure (could not reach the resident).
    Business outcomes are NOT exceptions — they live in the returned envelope's
    'success' field, same as the CLI's exit code."""
    def __init__(self, code, msg):
        super().__init__(f"[exit {code}] {msg}")
        self.code = code


# ---------------------------------------------------------------- pipe address
def _dotnet_tempdir():
    # Mirror .NET Path.GetTempPath() on Unix exactly: $TMPDIR else /tmp.
    return os.environ.get("TMPDIR") or "/tmp"


def _canonical_path(file_path):
    """Match the path officecli's resident hashes into the pipe name. On Windows
    it derives the name from GetFullPath, which expands 8.3 short components
    (RUNNER~1, or any user name > 8 chars under %TEMP%) to their long form.
    os.path.abspath does NOT expand 8.3, so a short path hashes to a different
    pipe and every connect fails with ENOENT — hence realpath, which does expand
    it. realpath needs the file to exist; fall back to the abspath when it
    doesn't. realpath ALSO resolves symlinks/junctions, which GetFullPath does
    not; harmless here because we hand this resolved path to the resident, so the
    server's GetFullPath sees the already-resolved string and both sides hash the
    same thing. Windows only — on unix officecli uses GetFullPath (no symlink
    resolution), so realpath would diverge there (e.g. /tmp -> /private/tmp on
    macOS)."""
    resolved = os.path.abspath(file_path)
    if _IS_WIN:
        try:
            return os.path.realpath(resolved)
        except OSError:
            pass
    return resolved


def pipe_paths(file_path):
    """(main, ping) pipe addresses for a document path. Exposed for debugging."""
    full = _canonical_path(file_path)
    if _IS_MAC or _IS_WIN:
        full = full.upper()                       # Linux: case-sensitive, no upper
    h = hashlib.sha256(full.encode("utf-8")).hexdigest().upper()[:16]
    name = f"officecli-{h}"
    if _IS_WIN:
        return rf"\\.\pipe\{name}", rf"\\.\pipe\{name}-ping"
    base = os.path.join(_dotnet_tempdir(), f"CoreFxPipe_{name}")
    return base, base + "-ping"


# ---------------------------------------------------------------- transport
# One attempt: bound the CONNECT, then block on the reply (no read timeout) —
# exactly like officecli's TrySend (Connect(timeout) + blocking PipeReadLine).
def _send_unix(sock_path, line, connect_timeout):
    s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    try:
        s.settimeout(connect_timeout)
        s.connect(sock_path)
        s.settimeout(None)                 # block on the reply; resident answers in turn
        s.sendall(line)
        buf = b""
        while not buf.endswith(b"\n"):
            chunk = s.recv(65536)
            if not chunk:
                break
            buf += chunk
        return buf
    finally:
        s.close()


def _send_win(pipe_path, line, connect_timeout):
    deadline = time.time() + connect_timeout
    while True:                            # bound the "open" (connect) phase
        try:
            f = _builtin_open(pipe_path, "r+b", buffering=0)   # not the module open()
            break
        except FileNotFoundError:
            # No pipe == no resident. Fail FAST, like _send_unix's connect()
            # raising ENOENT immediately — do NOT spin to the deadline. This is
            # what makes a max_retries=0 probe (_serves/alive) fail fast instead
            # of sitting through the whole connect_timeout when nothing is there.
            raise
        except OSError:
            # The pipe exists but the open lost the race (e.g. ERROR_PIPE_BUSY:
            # every server instance is mid-handoff). The resident IS alive, so
            # retry until the connect deadline.
            if time.time() > deadline:
                raise
            time.sleep(0.02)
    try:
        # FileIO.write (raw, buffering=0) issues a single WriteFile and may
        # return a short count, so loop until the whole request is out — a
        # truncated request leaves the resident blocking for a newline that
        # never comes, deadlocking the (untimed) reply read. Mirrors _send_unix's
        # sendall() and the C# client's Stream.Write.
        view = memoryview(line)
        sent = 0
        while sent < len(view):
            n = f.write(view[sent:])
            if n is None:                  # non-blocking handle not ready (shouldn't happen)
                continue
            sent += n
        buf = b""
        while not buf.endswith(b"\n"):     # blocking read, like PipeReadLine
            chunk = f.read(65536)
            if not chunk:
                break
            buf += chunk
        return buf
    finally:
        f.close()


def _rpc(sock_path, req, connect_timeout=_BUSY_CONNECT_TIMEOUT, max_retries=_BUSY_MAX_RETRIES):
    """Forward one request, mirroring officecli's TrySend: bounded connect + a few
    retries with backoff, then a blocking read. A retry only re-attempts the
    connect (before the command runs), so it never double-applies a mutation. If
    the command still can't be delivered, raise a busy/unresponsive error — never
    fall back to touching the file directly (that would race the resident).

    `max_retries` overrides the busy-retry count. Liveness probes (_serves) pass 0
    so a missing/stale pipe fails FAST instead of sleeping through ~0.3s of backoff
    — retrying a probe the resident isn't answering can't make it answer; the
    busy-retry policy is for delivering a real command to a slow-but-live pipe."""
    line = (json.dumps(req, ensure_ascii=False) + "\n").encode("utf-8")
    send = _send_win if _IS_WIN else _send_unix
    for attempt in range(max_retries + 1):
        try:
            raw = send(sock_path, line, connect_timeout)
            break
        except OSError as e:
            if attempt >= max_retries:
                raise OfficeCliError(-1,
                    f"resident is running but the command could not be delivered "
                    f"(pipe busy or unresponsive); retry, or close and reopen [{e}]")
            time.sleep(0.05 * (attempt + 1))    # = TrySend's 50*(n+1)ms backoff
    # utf-8-sig: the resident's StreamWriter (Encoding.UTF8) prepends a BOM the
    # C# StreamReader strips; we must too, or json.loads chokes on the leading .
    text = raw.decode("utf-8-sig")
    if not text.strip():
        # Empty/closed reply: the resident accepted the connection but closed
        # without a complete response (e.g. crashed mid-serve). We refuse to
        # re-send — the command may already have been APPLIED before the resident
        # died, so re-sending would double-apply a non-idempotent op — and raise
        # instead. officecli's TrySend now matches: its retry covers only the
        # connect phase (before the command is written); on an empty reply after a
        # successful write it returns null without re-sending, the C# equivalent of
        # this raise. _cmd's recovery then restarts a dead resident and retries once
        # (a fresh connect, before re-send), and _serves()/alive() (which swallow
        # OfficeCliError) read an empty reply as "not alive".
        raise OfficeCliError(-1,
            "resident closed the connection without a response "
            "(it may have crashed mid-command); retry, or close and reopen")
    return json.loads(text)


def _parse(resp):
    """Return the useful payload: the parsed JSON envelope (dict/list) if Stdout is
    a JSON object/array, otherwise the raw Stdout text ("" when empty). We accept
    ONLY dict/list from json.loads — a text-mode reply that happens to BE a bare
    JSON scalar ("42", "true", "null", a quoted string) must stay text, or the
    caller can't tell literal text "42" from the number 42 (and None from a missing
    key). Faithful to the response — no synthesizing a dict for view/raw text."""
    out = resp.get("Stdout", "")
    try:
        v = json.loads(out)
    except ValueError:
        return out
    return v if isinstance(v, (dict, list)) else out


def _strv(d):
    # Drop None-valued props (omit), matching how _cmd() drops None args — a prop
    # set to None means "don't send it", not "send empty string". Pass "" for
    # an explicit empty value.
    return {k: str(v) for k, v in d.items() if v is not None}


def _serves(ping_path, full_path, timeout=1.0):
    """Is a resident alive on `ping_path` AND serving `full_path`? Probes the
    always-responsive `-ping` pipe (officecli's TryConnect equivalent): it answers
    even while the MAIN pipe is busy. The path-match guards against a stale socket
    serving a different/renamed file. `full_path` must already be absolute.
    Single-shot (max_retries=0): a probe should fail fast, not sit through the
    busy-retry backoff that a real command delivery uses."""
    try:
        resp = _rpc(ping_path, {"Command": "__ping__"}, timeout, max_retries=0)
    except OfficeCliError:
        return False
    served = resp.get("Stdout", "").strip()   # ping echoes the served file path
    if not served:
        return False
    a = os.path.abspath(served)
    return a == full_path or ((_IS_MAC or _IS_WIN) and a.lower() == full_path.lower())


def _install_dir_candidate(name):
    """Where the official installer (install.sh / install.ps1) drops the binary:
    ~/.local/bin on macOS/Linux, %LOCALAPPDATA%\\OfficeCLI on Windows. Used only
    as a PATH-miss fallback (see _resolve_binary)."""
    if _IS_WIN:
        base = os.environ.get("LOCALAPPDATA")
        if not base:
            return None
        exe = name if name.lower().endswith(".exe") else name + ".exe"
        return os.path.join(base, "OfficeCLI", exe)
    return os.path.join(os.path.expanduser("~"), ".local", "bin", name)


def _resolve_binary(binary):
    """Resolve the officecli binary to invoke. Order: explicit path (a value with
    a path separator) is trusted as-is; otherwise a bare name is looked up on
    PATH; if PATH misses, fall back to the official installer's known location.

    Why the fallback: the installer adds its dir to PATH via the shell rc file, so
    a bare 'officecli' resolves in an interactive terminal — but NOT in processes
    that never sourced that rc (IDE-spawned Python, cron, systemd, CI). The binary
    is still sitting at the known install path; find it there instead of failing.

    Idempotent: an already-resolved absolute path passes straight through, so it's
    safe to call at every entry point (create + Document)."""
    if os.sep in binary or (os.altsep and os.altsep in binary):
        return binary                       # explicit path: trust the caller
    found = shutil.which(binary)
    if found:
        return found                        # on PATH: normal case
    cand = _install_dir_candidate(binary)   # PATH miss: try the known install dir
    if cand and os.path.isfile(cand) and os.access(cand, os.X_OK):
        return cand
    return binary                           # give up; _run_cli raises the helpful error


def _runs_ok(binary):
    """True iff `<binary> --version` actually runs and exits 0. Accept only a
    WORKING officecli — skip a present-but-broken file, and don't trigger a
    needless install when a usable officecli is already there."""
    try:
        return subprocess.run([binary, "--version"], capture_output=True).returncode == 0
    except OSError:
        return False


def _ensure_binary(binary, auto_install=True):
    """Resolve to a WORKING officecli, provisioning one if none is found and
    auto_install is set. An explicit path (with a separator) is trusted as-is;
    otherwise each candidate (PATH, then the installer's known location) is
    accepted only when `officecli --version` actually runs — so a present-but-
    broken binary is skipped and a usable one never triggers a needless install.
    install() picks install.sh (unix) or install.ps1 (Windows), so auto-install
    works on both."""
    if os.sep in binary or (os.altsep and os.altsep in binary):
        return binary                      # explicit path: trust the caller
    for cand in filter(None, (shutil.which(binary), _install_dir_candidate(binary))):
        if _runs_ok(cand):
            return cand                    # a working officecli is already here
    if auto_install:
        print("officecli CLI not found — installing from d.officecli.ai ...", file=sys.stderr)
        install()                          # CLI absent/unusable → official installer
        for cand in filter(None, (shutil.which(binary), _install_dir_candidate(binary))):
            if _runs_ok(cand):
                return cand
    return binary                          # give up; _run_cli raises the helpful error


def _run_cli(binary, argv):
    """Run `binary <argv...>` (capturing output). A missing binary surfaces as a
    clear OfficeCliError with install guidance, not a raw FileNotFoundError."""
    try:
        # text=True alone decodes with locale.getencoding() — the ANSI code page
        # on Windows (cp936/cp1252). The CLI always writes UTF-8, so a message
        # carrying a non-ASCII path raised UnicodeDecodeError (or mojibake) out
        # of subprocess instead of a clean OfficeCliError. The resident-pipe
        # route below already decodes utf-8; keep the two agreeing.
        return subprocess.run([binary, *argv], capture_output=True,
                              text=True, encoding="utf-8", errors="replace")
    except FileNotFoundError:
        raise OfficeCliError(127, _MISSING_CLI.format(bin=binary)) from None


# ---------------------------------------------------------------- the shell
class Document:
    def __init__(self, path, binary="officecli", timeout=30.0):
        # Canonical (Windows 8.3-expanded) so the pipe name AND the _serves()
        # path comparison both match what the resident reports.
        self.path = _canonical_path(path)
        self.bin = _resolve_binary(binary)
        self.timeout = timeout          # connect timeout (s); the reply read blocks
        self._main, self._ping = pipe_paths(self.path)
        self._restart_lock = threading.Lock()   # serialize dead-resident restarts
        self._start()

    def _start(self):
        # If a resident is ALREADY serving this file, reuse it — no process spawn.
        # Mirrors officecli, where a command after `create` reuses the resident
        # `create` auto-started instead of re-running `open`. _serves() is a real
        # liveness probe (ping the -ping pipe + verify the served path), not a
        # socket-file-exists check, so a stale/dead socket fails the probe and
        # falls through to `officecli open`, which replaces it via TryConnect.
        # (A plain os.path.exists() here would wrongly skip on a stale socket.)
        if _serves(self._ping, self.path):
            return
        # Otherwise spawn `officecli open` (one process). It's idempotent and uses
        # the same TryConnect to start a fresh resident or replace a stale socket.
        r = _run_cli(self.bin, ["open", self.path])
        if r.returncode != 0:
            raise OfficeCliError(r.returncode, r.stderr or r.stdout)

    # -- transport primitive: build {Command,Args,Props,Json}, forward, parse --
    def _cmd(self, command, args=None, props=None, as_json=True, timeout=None):
        # `as_json`, not `json`, so we don't shadow the imported json module.
        # timeout=None uses this Document's default (self.timeout). It bounds the
        # CONNECT/delivery (with retries); the reply read blocks, so a legitimately
        # slow command isn't cut off — it waits for the resident, like officecli.
        req = {"Command": command, "Json": as_json}
        if args:
            req["Args"] = {k: str(v) for k, v in args.items() if v is not None}
        if props is not None:
            req["Props"] = _strv(props)
        t = self.timeout if timeout is None else timeout
        try:
            return _rpc(self._main, req, t)
        except OfficeCliError:
            # Delivery failed after _rpc's own connect retries. Use the -ping pipe
            # to tell DEAD from BUSY — officecli's own distinction (alive()):
            #   • ALIVE but main pipe unresponsive → do NOT bypass it. officecli
            #     deliberately dropped the direct-file fallback: a second writer
            #     racing the live resident loses data on its eventual save. Re-raise
            #     the busy error so the caller can retry or close+reopen.
            #   • DEAD (crashed / stale socket) → restart with one `officecli open`
            #     and retry ONCE. Safe across reads and mutations: mutations live in
            #     memory until save/close, so a crash loses them and disk holds the
            #     last save — replaying against the restarted (disk-state) resident
            #     reproduces the lost op once, with nothing live to double-apply.
            if self.alive():
                raise
            # Serialize the restart across threads sharing this Document. Without
            # the lock, N concurrent callers each see alive()==False and each spawn
            # `officecli open`, leaving N-1 orphaned residents on the same file
            # (which can then race each other's save). Re-check alive() inside the
            # lock so only the first thread restarts; the rest find it back up.
            with self._restart_lock:
                if not self.alive():
                    self._start()
            return _rpc(self._main, req, t)

    # -- the surface: send ONE batch-shaped command, or a LIST of them ---------
    def send(self, item, as_json=True, timeout=None):
        """Forward ONE command in officecli's batch-item shape and return its
        parsed result (the JSON envelope, or raw text for content commands).

        `item` is exactly a dict you'd put in a `batch` list, e.g.
            {"command": "set", "path": "/Sheet1/A1", "props": {"text": "hi"}}
            {"command": "get", "path": "/Sheet1/A1"}
        Keys are officecli's batch fields; `command` (or `op`) picks the command,
        `props` becomes the property map, and every other key is forwarded
        verbatim as a command argument — no field list maintained here, so new
        officecli fields work without touching this shell.

        `as_json=False` requests plain-text output (view/raw/dump), mirroring the
        CLI's --json toggle."""
        command = item.get("command") or item.get("op")
        if not command:
            raise OfficeCliError(-1, "send(item): item needs a 'command' (or 'op') key")
        args = {k: v for k, v in item.items() if k not in ("command", "op", "props")}
        return _parse(self._cmd(command, args, item.get("props"),
                                as_json=as_json, timeout=timeout))

    def batch(self, items, force=True, stop_on_error=False, timeout=None):
        """Forward officecli's `batch` command: apply a LIST of the same item
        dicts as `send` in ONE round-trip — the fast path for many writes. Same
        contract as `send`, just plural."""
        args = {"batchJson": json.dumps(items, ensure_ascii=False),
                "force": force, "stopOnError": stop_on_error}
        return _parse(self._cmd("batch", args, timeout=timeout))

    def _set_idle_timeout(self, seconds):
        # Best-effort idle-timeout upgrade, served on the always-responsive ping
        # pipe (bypasses _commandLock, answers even while the main pipe is busy).
        # Mirrors ResidentClient.SendSetIdleTimeout: a failure is non-fatal — the
        # resident is still usable, it just keeps its original idle schedule.
        # Single-shot (max_retries=0): don't sit through the busy backoff for a
        # best-effort nicety.
        try:
            _rpc(self._ping, {"Command": "__set-idle-timeout__",
                              "Args": {"seconds": str(seconds)}},
                 self.timeout, max_retries=0)
        except OfficeCliError:
            pass

    def alive(self, timeout=1.0):
        """Return True iff a resident is alive AND serving this file. Probes the
        always-responsive `-ping` pipe (officecli's TryConnect), which answers even
        while the MAIN pipe is busy — so it distinguishes "alive but busy" from
        "gone". This is the discriminator `_cmd` uses on a delivery failure (busy →
        raise, gone → restart+retry); send/batch already auto-recover from a gone
        resident, so call this only when you want to check liveness yourself."""
        return _serves(self._ping, self.path, timeout)

    # -- lifecycle ------------------------------------------------------------
    def close(self):
        # = `officecli close`: stop the resident. It flushes the in-memory doc to
        # disk as it shuts down (handler.Dispose), so no separate save is needed —
        # verified: a set followed by __close__ alone lands on disk.
        #
        # The resident acks AFTER shutting down, so a missing/empty ack (lost to a
        # crash or the 5s write-timeout) still means "closed". A real shutdown
        # data-loss is a NON-empty error response, so it surfaces through _parse.
        try:
            return _parse(_rpc(self._ping, {"Command": "__close__"}, self.timeout))
        except OfficeCliError:
            # Only swallow if the resident is actually gone. If it's still alive
            # (ping pipe was momentarily unreachable/busy), the close did NOT take
            # effect — re-raise, or the caller wrongly believes the file is released
            # and may race a re-open/overwrite.
            if self.alive():
                raise
            return ""   # resident gone / ack lost — end state is "closed"

    def __enter__(self):
        return self

    def __exit__(self, *a):
        # `with` means "I manage this session" → close on exit. To only borrow a
        # resident another program owns, DON'T use `with` and DON'T call close():
        #     d = officecli.open(f); d.send(...)              # left running
        self.close()


def create(path, *args, binary="officecli", timeout=30.0, auto_install=True):
    """Create a blank Office document and return a live `Document` handle for it.

    Parallel to `open`: both return the session handle you actually work with —
    they differ only in the file's expected state. `open` requires an existing
    file; `create` makes a new one (like file mode "x" vs "r"). Extra CLI flags
    pass through verbatim, so there's no option list maintained here:
        with officecli.create("report.xlsx", "--force") as doc:
            doc.send({"command": "set", "path": "/Sheet1/A1", "props": {"text": "hi"}})
        officecli.create("doc", "--type", "docx")

    One CLI spawn (`officecli create`), which also auto-starts a resident for the
    new file; the returned Document binds to THAT resident (no second spawn).
    Raises OfficeCliError on failure, inheriting officecli's exact semantics:
      • file held by a LIVE resident → file_locked (close it first). We do NOT
        silently close+overwrite it — in a shared workspace that resident may be
        another owner's active session.
      • file exists without --force → file_exists (pass "--force" to overwrite)."""
    full = os.path.abspath(path)
    binary = _ensure_binary(binary, auto_install)
    r = _run_cli(binary, ["create", full, *args])
    if r.returncode != 0:
        raise OfficeCliError(r.returncode, r.stderr or r.stdout)
    # create auto-started a resident for the new file; bind a handle to it
    # (Document.__init__ -> _start -> _serves finds it alive, so no extra spawn).
    return Document(full, binary=binary, timeout=timeout)


def open(path, binary="officecli", timeout=30.0, auto_install=True):
    """Open an EXISTING document and return a live `Document` handle (parallel to
    `create`, which makes a new file). `officecli open` is idempotent: it reuses a
    resident already serving this file or starts one — and if a live resident is
    already up, no process is spawned at all.

    Lifecycle:
      Owner  — `with officecli.open(f) as d: ...`   (exit closes the resident)
      Borrow — `d = officecli.open(f); d.send(...)` (no `with`/close → left running)

    Failure model (applies to every send/batch on the handle):
      • resident DEAD/gone (crash, idle-timeout, missing pipe) → transparently
        restarted and the command retried once; the caller sees no error.
      • resident ALIVE but the pipe is unresponsive (busy) → raises OfficeCliError
        — never a deadlock, and never bypassing the live resident (that would race
        its save and lose data). Retry, or close() and reopen.

    `timeout` bounds command DELIVERY (connect + retries) in seconds, mirroring
    officecli's TrySend; the reply read itself blocks (a busy resident answers in
    turn). Override per call via send(..., timeout=...) / batch(..., timeout=...);
    use alive() to probe liveness."""
    doc = Document(path, binary=_ensure_binary(binary, auto_install), timeout=timeout)
    # Mirror CLI `open`: when reusing a resident `create` auto-started with a
    # short 60s timeout, upgrade it to the 12min interactive window. (If _start
    # spawned `officecli open` instead, that path already set 12min; re-sending
    # is idempotent.)
    doc._set_idle_timeout(_OPEN_IDLE_SECONDS)
    return doc


def install():
    """Install the officecli CLI binary via its OFFICIAL installer — install.sh on
    unix, install.ps1 on Windows. Reuses officecli's own installers (platform
    detection + checksum + ~/.local/bin or %LOCALAPPDATA%\\OfficeCLI), rather than
    reimplementing download logic that would drift from upstream.

    Called automatically by open()/create() when the CLI is missing (pass
    auto_install=False to disable), and exposed directly as `python -m officecli
    install`. Returns None on success; raises OfficeCliError on failure. Output is
    NOT captured, so the installer's progress and checksum lines stream to the
    user."""
    if _IS_WIN:
        print(f"Installing officecli via {_INSTALL_PS1_MIRROR} (github fallback) ...", file=sys.stderr)
        # Windows PowerShell (powershell.exe) ships with the OS; -ExecutionPolicy
        # Bypass lets the remote script run without changing machine policy. Fetch
        # the script mirror-first, github fallback, then run it.
        ps = (f"$s = try {{ irm '{_INSTALL_PS1_MIRROR}' }} "
              f"catch {{ irm '{_INSTALL_PS1_GITHUB}' }}; $s | iex")
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps])
        if r.returncode != 0:
            raise OfficeCliError(r.returncode,
                f"officecli install failed (exit {r.returncode}). Run manually:\n"
                f"    irm {_INSTALL_PS1_MIRROR} | iex")
        return None
    print(f"Installing officecli via {_INSTALL_SH_MIRROR} (github fallback) ...", file=sys.stderr)
    # (curl mirror || curl github) | bash — the subshell emits whichever fetch
    # succeeds; the group keeps the pipe bound to the whole fallback. Output is
    # NOT captured, so progress and checksum lines stream to the user.
    sh = f"(curl -fsSL {_INSTALL_SH_MIRROR} 2>/dev/null || curl -fsSL {_INSTALL_SH_GITHUB}) | bash"
    r = subprocess.run(["bash", "-c", sh])
    if r.returncode != 0:
        raise OfficeCliError(r.returncode,
            f"officecli install failed (exit {r.returncode}). Run manually:\n"
            f"    curl -fsSL {_INSTALL_SH_MIRROR} | bash")
    return None


# Advertised surface = the command shell + its error. pipe_paths stays importable
# (officecli.pipe_paths) as a debug helper but isn't part of the command API.
__all__ = ["open", "create", "install", "Document", "OfficeCliError"]


if __name__ == "__main__":
    # `python -m officecli install` — bootstrap the CLI binary.
    if len(sys.argv) >= 2 and sys.argv[1] == "install":
        install()
    else:
        print("usage: python -m officecli install", file=sys.stderr)
        sys.exit(2)
