// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    public static RootCommand BuildRootCommand()
    {
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON (AI-friendly)" };

        var rootCommand = new RootCommand("""
            officecli: AI-friendly CLI for Office documents (.docx, .xlsx, .pptx)

            Run 'officecli help' for the schema-driven capability reference (formats, elements, properties).
            See the Commands section below for the full list of subcommands.
            """);
        rootCommand.Add(jsonOption);

        // ==================== open command (start resident) ====================
        var openFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var openCommand = new Command("open", "Start a resident process to keep the document in memory for faster subsequent commands");
        openCommand.Add(openFileArg);
        openCommand.Add(jsonOption);

        openCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(openFileArg)!;
            var filePath = file.FullName;

            // If already running, reuse the existing resident. This covers
            // two cases with the same code path:
            //   (a) user previously called `open` explicitly, or
            //   (b) `create` just auto-started a short-lived (60s) resident.
            // In either case we upgrade the idle timeout to the default 12min
            // via the __set-idle-timeout__ ping RPC. Failure is non-fatal —
            // the resident is still usable, it'll just exit on its original
            // schedule. `open` is idempotent, so repeated calls are safe.
            const int DefaultOpenIdleSeconds = 12 * 60;
            if (ResidentClient.TryConnect(filePath, out _))
            {
                ResidentClient.SendSetIdleTimeout(filePath, DefaultOpenIdleSeconds);
                var msg = $"Opened {file.Name} (reusing running resident, idle timeout set to 12min). "
                        + $"Still pass the file path on every command (e.g. get \"{file.Name}\" /body); run 'close {file.Name}' when done.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
                return 0;
            }

            if (!TryStartResidentProcess(filePath, idleSeconds: null, out var startError))
                throw new InvalidOperationException(startError);

            var startedMsg = $"Opened {file.Name} (resident started). "
                           + $"Still pass the file path on every command (e.g. get \"{file.Name}\" /body); run 'close {file.Name}' when done.";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(startedMsg));
            else Console.WriteLine(startedMsg);
            return 0;
        }, json); });

        rootCommand.Add(openCommand);

        // ==================== close command (stop resident) ====================
        var closeFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var closeCommand = new Command("close", "Flush in-memory changes to disk and stop the resident (releases the file). Use 'save' instead to flush but keep the resident warm. Either is needed before a non-officecli program reads the file; a live resident also auto-flushes shortly after going idle (adaptive 2-10s; see OFFICECLI_RESIDENT_FLUSH: each|auto|<seconds>|off).");
        closeCommand.Add(closeFileArg);
        closeCommand.Add(jsonOption);

        closeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(closeFileArg)!;
            if (ResidentClient.SendCloseWithResponse(file.FullName, out var closeResp))
            {
                // BUG-BT-R26-2: resident may report a non-zero shutdown
                // (e.g. file vanished mid-session → data loss). Bubble
                // that up instead of pretending the close succeeded.
                if (closeResp != null && closeResp.ExitCode != 0)
                {
                    var err = !string.IsNullOrEmpty(closeResp.Stderr)
                        ? closeResp.Stderr
                        : $"Resident close reported error (exit {closeResp.ExitCode})";
                    throw new InvalidOperationException(err);
                }
                // BUG-INTERVIEW-EDIT-R10-B: resident reports advisory warnings
                // (e.g. backing file missing at original path) via Stderr with
                // exit=0. Forward to the client's stderr so the user sees the
                // warning instead of a silent success.
                if (closeResp != null && !string.IsNullOrEmpty(closeResp.Stderr))
                    Console.Error.WriteLine(closeResp.Stderr);
                var msg = $"Resident closed for {file.Name}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
            }
            else
            {
                // No resident is holding this file. In the non-resident model
                // every mutation already eager-saved to disk, so there is
                // nothing to flush or shut down — treat close as an idempotent
                // no-op SUCCESS, not an error. This lets "edit, then close when
                // done" be a safe habit regardless of whether a resident was
                // ever started; erroring here used to actively discourage it.
                var msg = $"{file.Name} is already saved to disk; nothing to close.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
            }
            return 0;
        }, json); });

        rootCommand.Add(closeCommand);

        // ==================== __resident-serve__ (internal, hidden) ====================
        var serveFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var serveCommand = new Command("__resident-serve__", "Internal: run resident server (do not call directly)");
        serveCommand.Hidden = true;
        serveCommand.Add(serveFileArg);

        serveCommand.SetAction(result =>
        {
            var file = result.GetValue(serveFileArg)!;
            // Per-file singleton guard. TryResident's probe-then-spawn has an
            // inherent race: N clients probing an un-owned file concurrently
            // all fail the ping and all spawn a resident. Each spawned server
            // held its own full in-memory copy and whole-file-overwrote on
            // flush — concurrent writers silently lost every edit except the
            // last flusher's (observed: 40 parallel sets → 0-2 cells on
            // disk, all reporting success). Acquire an exclusive lock file
            // BEFORE opening the document; losers exit quietly and their
            // clients reconnect to the winner via the re-probe in
            // TryResident.
            FileStream? residentLock = null;
            var lockPath = Path.Combine(Path.GetTempPath(),
                ResidentServer.GetPipeName(file.FullName) + ".lock");
            for (int attempt = 0; attempt < 3 && residentLock == null; attempt++)
            {
                try
                {
                    residentLock = new FileStream(lockPath, FileMode.OpenOrCreate,
                        FileAccess.ReadWrite, FileShare.None,
                        bufferSize: 1, FileOptions.DeleteOnClose);
                }
                catch (IOException)
                {
                    // Another resident holds (or is acquiring) the lock. If it
                    // is already serving, we're redundant — exit and let the
                    // client reconnect. Brief retry covers the window where
                    // the winner crashed without deleting the lock.
                    if (ResidentClient.TryConnect(file.FullName, out var winnerPipe)) return;
                    Thread.Sleep(150);
                }
            }
            if (residentLock == null) return;
            using var heldLock = residentLock;
            using var server = new ResidentServer(file.FullName);
            server.RunAsync().GetAwaiter().GetResult();
        });

        rootCommand.Add(serveCommand);

        // Register commands from partial files
        rootCommand.Add(BuildWatchCommand(jsonOption));
        rootCommand.Add(BuildUnwatchCommand());
        // BC aliases — mark/unmark/get-marks/goto were promoted to `watch <sub>`
        // subcommands; the top-level forms are kept registered but hidden so
        // existing scripts and tests keep working. Remove after a deprecation
        // window once external usage has migrated.
        var markBc = BuildMarkCommand(jsonOption);          markBc.Hidden = true; rootCommand.Add(markBc);
        var unmarkBc = BuildUnmarkMarkCommand(jsonOption);  unmarkBc.Hidden = true; rootCommand.Add(unmarkBc);
        var getMarksBc = BuildGetMarksCommand(jsonOption);  getMarksBc.Hidden = true; rootCommand.Add(getMarksBc);
        var gotoBc = BuildGotoCommand(jsonOption);          gotoBc.Hidden = true; rootCommand.Add(gotoBc);
        rootCommand.Add(BuildViewCommand(jsonOption));
        rootCommand.Add(BuildGetCommand(jsonOption));
        rootCommand.Add(BuildQueryCommand(jsonOption));
        rootCommand.Add(BuildSetCommand(jsonOption));
        rootCommand.Add(BuildAddCommand(jsonOption));
        rootCommand.Add(BuildRemoveCommand(jsonOption));
        rootCommand.Add(BuildMoveCommand(jsonOption));
        rootCommand.Add(BuildSwapCommand(jsonOption));
        rootCommand.Add(BuildRefreshCommand(jsonOption));
        rootCommand.Add(BuildRawCommand(jsonOption));
        rootCommand.Add(BuildRawSetCommand(jsonOption));
        rootCommand.Add(BuildAddPartCommand(jsonOption));
        rootCommand.Add(BuildValidateCommand(jsonOption));
        rootCommand.Add(BuildSaveCommand(jsonOption));
        rootCommand.Add(BuildBatchCommand(jsonOption));
        rootCommand.Add(BuildDumpCommand(jsonOption));
        rootCommand.Add(BuildImportCommand(jsonOption));
        rootCommand.Add(BuildCreateCommand(jsonOption));
        rootCommand.Add(BuildMergeCommand(jsonOption));
        rootCommand.Add(BuildPluginsCommand(jsonOption));
        rootCommand.Add(BuildDeckCommand(jsonOption));

        foreach (var stub in BuildIntegrationStubCommands())
            rootCommand.Add(stub);

        rootCommand.Add(BuildHelpCommand(jsonOption, rootCommand));

        return rootCommand;
    }

    // ==================== Helper: fork a __resident-serve__ subprocess ====================
    //
    // Used by both `open` (explicit) and `create` (auto-start after
    // creating a blank file). Forks the current executable with the
    // internal __resident-serve__ verb and waits up to 5s for the ping
    // pipe to respond, so callers get a definitive success/fail answer.
    //
    // `idleSeconds` overrides the child's idle-exit timeout via the
    // OFFICECLI_RESIDENT_IDLE_SECONDS env var (1..86400). Passing null
    // inherits the server default (12 minutes). `create` passes 60 so
    // an auto-started resident that nobody follows up on exits quickly.
    //
    // Caller must first verify no resident is already running for this
    // file (e.g. via ResidentClient.TryConnect) — this helper always
    // starts a fresh child.
    internal static bool TryStartResidentProcess(string filePath, int? idleSeconds, out string? error)
    {
        error = null;
        var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (exePath == null)
        {
            error = "Cannot determine executable path.";
            return false;
        }

        // The resident is a long-lived background server that talks to clients
        // only over a named pipe — it must inherit NOTHING from the transient
        // CLI invocation that spawns it. The failure this guards against: on
        // Windows, .NET's UseShellExecute=false path calls CreateProcess with
        // bInheritHandles=TRUE, which duplicates EVERY inheritable handle in
        // our process into the child — including any handle to the caller's
        // stdout/stderr pipe. When the caller's stdout is a pipe ($(), | cat,
        // CI, an SDK/agent shell), that leaked write handle keeps the pipe open
        // in the resident, so the caller's read never sees EOF until the
        // resident idle-exits (~60s) even though the command already returned.
        //
        // Clearing the inherit flag on our three std handles is NOT enough: if
        // a second inheritable handle to the same pipe exists in our process
        // (a duplicate left by an injected module, the runtime, or the
        // launching shell), it still leaks. The robust fix is to inherit ONLY
        // the handles we explicitly hand the child: CreateProcess with a
        // PROC_THREAD_ATTRIBUTE_HANDLE_LIST whitelist (the child's own std
        // handles), so no stray handle can cross no matter how many exist.
        //
        // On macOS/Linux, posix_spawn inherits fds unless the child's
        // stdout/stderr are explicitly redirected. RedirectStandardOutput /
        // RedirectStandardError = true makes .NET plumb a fresh pipe from
        // parent to child, so the caller's shell pipe (e.g. `| tail -1`,
        // $(...)) is NOT inherited and EOFs promptly when the client exits.
        // See ResidentStdoutInheritanceTests for the regression lock-in.
        bool isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

        // Uniform view over the two spawn paths so the readiness loop below is
        // identical: has the child exited yet, read its stderr once (crash
        // diagnostics), release our handle to it.
        Func<bool> hasExited;
        Func<string> readStderr;
        Action dispose;

        if (isWindows)
        {
            if (!StartResidentWindows(exePath, filePath, idleSeconds, out var hProcess, out readStderr, out var startError))
            {
                error = startError ?? "Failed to start resident process.";
                return false;
            }
            hasExited = () => WaitForSingleObject(hProcess, 0) == 0 /* WAIT_OBJECT_0 */;
            dispose = () => CloseHandle(hProcess);
        }
        else
        {
            // CONSISTENCY(child-process-args): forward verb + path via ArgumentList,
            // not a hand-quoted Arguments string. .NET re-parses the Arguments
            // string with Windows-style quoting rules even on Unix, so a filePath
            // containing a literal '"' (legal on macOS/Linux) or a trailing '\'
            // would split into stray argv and the resident would reject startup.
            // ArgumentList passes argv losslessly. Matches BlankDocCreator /
            // FormatHandlerSession, which fork this same exe the same way.
            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                ArgumentList = { "__resident-serve__", filePath },
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8
            };
            if (idleSeconds.HasValue)
                startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = idleSeconds.Value.ToString();

            Process? process = Process.Start(startInfo);
            if (process == null)
            {
                error = "Failed to start resident process.";
                return false;
            }
            hasExited = () => process.HasExited;
            readStderr = () => process.StandardError.ReadToEnd();
            dispose = () => process.Dispose();
        }

        // Wait briefly for the server to start accepting connections.
        for (int i = 0; i < 50; i++) // up to 5 seconds
        {
            Thread.Sleep(100);
            if (ResidentClient.TryConnect(filePath, out _))
            {
                dispose();
                return true;
            }
            if (hasExited())
            {
                var stderr = readStderr();
                // CONSISTENCY(cli-error-first-line): the resident process dumps its
                // full call stack on a startup crash; surface only the first line
                // (typically the exception message). The stack is still in the
                // resident's own log if needed for diagnostics — keeping it out
                // of the user-facing CLI error avoids burying the actual cause.
                var firstLine = string.IsNullOrEmpty(stderr)
                    ? ""
                    : stderr.Split('\n', 2)[0].TrimEnd('\r').Trim();
                error = string.IsNullOrEmpty(firstLine)
                    ? "Resident process exited."
                    : $"Resident process exited. {firstLine}";
                dispose();
                return false;
            }
        }

        error = "Resident process started but not responding.";
        dispose();
        return false;
    }

    // ==================== Win32 resident spawn (Windows) ====================
    //
    // Spawn __resident-serve__ so it inherits ONLY the child's own std handles
    // (stdin/stdout -> NUL, stderr -> a private pipe we read on a startup
    // crash), never the caller's console/pipe handles. The explicit handle
    // whitelist means no stray inheritable handle can cross into the resident,
    // regardless of how many exist. UseShellExecute stays false, so args go via
    // a CommandLineToArgvW-safe quoted string and the idle override reaches the
    // child through the inherited environment (no window, works headless).
    private static bool StartResidentWindows(string exePath, string filePath, int? idleSeconds,
        out nint hProcess, out Func<string> readStderr, out string? error)
    {
        hProcess = 0;
        readStderr = static () => "";
        error = null;

        var sa = new SECURITY_ATTRIBUTES
        {
            nLength = Marshal.SizeOf<SECURITY_ATTRIBUTES>(),
            lpSecurityDescriptor = 0,
            bInheritHandle = 1
        };

        nint nulIn = CreateFileW("NUL", GENERIC_READ, FILE_SHARE_RW, ref sa, OPEN_EXISTING, 0, 0);
        nint nulOut = CreateFileW("NUL", GENERIC_WRITE, FILE_SHARE_RW, ref sa, OPEN_EXISTING, 0, 0);
        if (!CreatePipe(out nint errRead, out nint errWrite, ref sa, 0))
        {
            error = "CreatePipe failed: " + Marshal.GetLastWin32Error();
            return false;
        }
        SetHandleInformation(errRead, HANDLE_FLAG_INHERIT, 0); // read end stays private to us

        var siex = new STARTUPINFOEX();
        siex.StartupInfo.cb = Marshal.SizeOf<STARTUPINFOEX>();
        siex.StartupInfo.dwFlags = (int)STARTF_USESTDHANDLES;
        siex.StartupInfo.hStdInput = nulIn;
        siex.StartupInfo.hStdOutput = nulOut;
        siex.StartupInfo.hStdError = errWrite;

        nint size = 0;
        InitializeProcThreadAttributeList(0, 1, 0, ref size);
        nint attr = Marshal.AllocHGlobal(size);
        var inheritList = new[] { nulIn, nulOut, errWrite };
        var pin = GCHandle.Alloc(inheritList, GCHandleType.Pinned);
        bool ok = false;
        try
        {
            if (!InitializeProcThreadAttributeList(attr, 1, 0, ref size))
            {
                error = "InitializeProcThreadAttributeList failed: " + Marshal.GetLastWin32Error();
                return false;
            }
            if (!UpdateProcThreadAttribute(attr, 0, PROC_THREAD_ATTRIBUTE_HANDLE_LIST,
                    pin.AddrOfPinnedObject(), (nint)(nint.Size * inheritList.Length), 0, 0))
            {
                error = "UpdateProcThreadAttribute failed: " + Marshal.GetLastWin32Error();
                return false;
            }
            siex.lpAttributeList = attr;

            var cmd = new StringBuilder();
            cmd.Append(EscapeWindowsArg(exePath)).Append(' ')
               .Append(EscapeWindowsArg("__resident-serve__")).Append(' ')
               .Append(EscapeWindowsArg(filePath));

            // The resident reads OFFICECLI_RESIDENT_IDLE_SECONDS from its
            // environment; with bInheritHandles handled explicitly we pass a
            // null env block (child inherits ours). Set the override only around
            // the spawn so we don't mutate our own process environment for good.
            string? prevIdle = null;
            bool setIdle = idleSeconds.HasValue;
            if (setIdle)
            {
                prevIdle = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_IDLE_SECONDS");
                Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_IDLE_SECONDS", idleSeconds!.Value.ToString());
            }
            try
            {
                ok = CreateProcessW(null, cmd, 0, 0, /*bInheritHandles*/ true,
                    CREATE_NO_WINDOW | EXTENDED_STARTUPINFO_PRESENT, 0, null, ref siex, out var pi);
                if (ok)
                {
                    CloseHandle(pi.hThread);
                    hProcess = pi.hProcess;
                }
            }
            finally
            {
                if (setIdle) Environment.SetEnvironmentVariable("OFFICECLI_RESIDENT_IDLE_SECONDS", prevIdle);
            }
            if (!ok)
            {
                error = "CreateProcess failed: " + Marshal.GetLastWin32Error();
                return false;
            }

            nint errReadCaptured = errRead;
            readStderr = () =>
            {
                try
                {
                    var sb = new StringBuilder();
                    var buf = new byte[4096];
                    while (ReadFile(errReadCaptured, buf, (uint)buf.Length, out uint n, 0) && n > 0)
                        sb.Append(Encoding.UTF8.GetString(buf, 0, (int)n));
                    return sb.ToString();
                }
                catch { return ""; }
            };
            return true;
        }
        finally
        {
            if (attr != 0) { DeleteProcThreadAttributeList(attr); Marshal.FreeHGlobal(attr); }
            pin.Free();
            // Our copies of the child's inheritable handles; the child holds its
            // own inherited copies. errRead stays open for readStderr (released
            // when this short-lived CLI process exits) unless the spawn failed.
            CloseHandle(nulIn);
            CloseHandle(nulOut);
            CloseHandle(errWrite);
            if (!ok) CloseHandle(errRead);
        }
    }

    /// <summary>
    /// Quote one argv token for the Windows command line so CommandLineToArgvW
    /// round-trips it exactly (spaces, quotes, trailing backslashes).
    /// </summary>
    private static string EscapeWindowsArg(string arg)
    {
        if (arg.Length > 0 && arg.IndexOfAny(new[] { ' ', '\t', '"' }) < 0)
            return arg;
        var sb = new StringBuilder();
        sb.Append('"');
        int slashes = 0;
        foreach (char c in arg)
        {
            if (c == '\\') { slashes++; continue; }
            if (c == '"') { sb.Append('\\', slashes * 2 + 1); sb.Append('"'); }
            else { sb.Append('\\', slashes); sb.Append(c); }
            slashes = 0;
        }
        sb.Append('\\', slashes * 2);
        sb.Append('"');
        return sb.ToString();
    }

    // ==================== Win32 P/Invoke ====================

    private const uint HANDLE_FLAG_INHERIT = 0x00000001;
    private const uint STARTF_USESTDHANDLES = 0x00000100;
    private const uint CREATE_NO_WINDOW = 0x08000000;
    private const uint EXTENDED_STARTUPINFO_PRESENT = 0x00080000;
    private static readonly nint PROC_THREAD_ATTRIBUTE_HANDLE_LIST = 0x00020002;
    private const uint GENERIC_READ = 0x80000000;
    private const uint GENERIC_WRITE = 0x40000000;
    private const uint FILE_SHARE_RW = 0x00000003;
    private const uint OPEN_EXISTING = 3;

    [StructLayout(LayoutKind.Sequential)]
    private struct SECURITY_ATTRIBUTES { public int nLength; public nint lpSecurityDescriptor; public int bInheritHandle; }

    [StructLayout(LayoutKind.Sequential)]
    private struct STARTUPINFO
    {
        public int cb; public nint lpReserved, lpDesktop, lpTitle;
        public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public short wShowWindow, cbReserved2; public nint lpReserved2, hStdInput, hStdOutput, hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct STARTUPINFOEX { public STARTUPINFO StartupInfo; public nint lpAttributeList; }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_INFORMATION { public nint hProcess, hThread; public int dwProcessId, dwThreadId; }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetHandleInformation(nint hObject, uint dwMask, uint dwFlags);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(nint hObject);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern nint CreateFileW(string name, uint access, uint share, ref SECURITY_ATTRIBUTES sa, uint disposition, uint flags, nint template);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreatePipe(out nint hReadPipe, out nint hWritePipe, ref SECURITY_ATTRIBUTES sa, uint size);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ReadFile(nint hFile, byte[] buffer, uint count, out uint read, nint overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(nint handle, uint ms);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateProcessW(string? applicationName, StringBuilder commandLine,
        nint processAttributes, nint threadAttributes, bool inheritHandles, uint creationFlags,
        nint environment, string? currentDirectory, ref STARTUPINFOEX startupInfo, out PROCESS_INFORMATION processInformation);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool InitializeProcThreadAttributeList(nint list, int count, int flags, ref nint size);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UpdateProcThreadAttribute(nint list, uint flags, nint attribute, nint value, nint size, nint previous, nint returnSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern void DeleteProcThreadAttributeList(nint list);

    // ==================== Helper: try forwarding to resident ====================
    //
    // Two-step protocol (CONSISTENCY(resident-two-step): same shape as
    // CommandBuilder.Batch.cs's resident branch):
    //   1. Ping-pipe probe via TryConnect — fast (100ms) and isolated from the
    //      main command queue, so it stays responsive even under flood. Tells
    //      us definitively whether a resident owns this file.
    //   2. If yes, send the command on the main pipe with a generous connect
    //      timeout + a few retries. If the send STILL fails, surface a
    //      distinct "busy" error (exit code 3) instead of falling back to
    //      DocumentHandlerFactory.Open — the old silent fallback could race
    //      the live resident and lose writes.
    //   3. If no resident, return null so the caller opens the file directly.
    //
    // Exit code 3 is reserved for "resident is alive but couldn't deliver the
    // command" so callers can distinguish it from a command-level failure.
    private const int ResidentBusyExitCode = 3;
    private const int ResidentBusyConnectTimeoutMs = 30000;
    private const int ResidentBusyMaxRetries = 3;

    internal static int? TryResident(string filePath, Action<ResidentRequest> configure, bool json = false)
    {
        // Step 1: does a resident own this file? Probe via the -ping pipe,
        // which is never serialized behind main-pipe commands.
        if (!ResidentClient.TryConnect(filePath, out _))
        {
            // No resident running — auto-start one to avoid file-lock conflicts
            // when multiple commands hit the same file in parallel.
            // Opt-out: OFFICECLI_NO_AUTO_RESIDENT=1 disables auto-start (e.g.
            // sandbox environments where named pipes may not work reliably).
            var noAuto = Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT");
            if (noAuto == "1" || string.Equals(noAuto, "true", StringComparison.OrdinalIgnoreCase))
                return null;

            if (!TryStartResidentProcess(filePath, idleSeconds: 60, out _))
            {
                // Startup failed — maybe another process just started a resident
                // for the same file (parallel race). Re-probe before giving up.
                if (!ResidentClient.TryConnect(filePath, out _))
                    return null; // truly no resident → caller falls back to direct file access
            }
            // Intentionally no user-facing hint here. UX testing with an AI
            // agent showed a standalone "background process" hint on a random
            // mid-batch command (e.g. `get`) creates low-grade anxiety without
            // giving the caller a concrete action — auto-close in 60s already
            // handles the cleanup, and other officecli commands work normally
            // through the resident regardless. The `create` command keeps a
            // small inline suffix on its success line because it's contextual
            // to a freshly-created file, not a nag fired from anywhere.
        }

        var request = new ResidentRequest();
        configure(request);
        if (json) request.Json = true;

        // Step 2: resident is confirmed alive — wait for our turn in the main
        // pipe queue. Do NOT silently fall back on failure; letting a second
        // writer touch the file while the resident holds it in memory loses
        // data on the resident's eventual save.
        var response = ResidentClient.TrySend(
            filePath, request,
            maxRetries: ResidentBusyMaxRetries,
            connectTimeoutMs: ResidentBusyConnectTimeoutMs);

        if (response == null)
        {
            var fileName = Path.GetFileName(filePath);
            var msg = $"Resident for {fileName} is running but the command could not be delivered (main pipe busy or unresponsive). Retry, or run 'officecli close {fileName}' and try again.";
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
            else
                Console.Error.WriteLine($"Error: {msg}");
            return ResidentBusyExitCode;
        }

        if (json)
        {
            // JSON mode: resident already built the envelope, just pass through
            if (!string.IsNullOrEmpty(response.Stdout))
                Console.WriteLine(response.Stdout);
        }
        else
        {
            if (!string.IsNullOrEmpty(response.Stdout))
                Console.WriteLine(response.Stdout);
            if (!string.IsNullOrEmpty(response.Stderr))
                Console.Error.WriteLine(response.Stderr);
        }

        return response.ExitCode;
    }


    // ContainsNullByte — defensive guard for batch input. OOXML / xml-1.0
    // forbids U+0000 in any element or attribute content; an unfiltered NUL
    // reaches the SDK's xml writer at save time and throws an XmlException
    // that aborts the in-flight session and corrupts the resident's view.
    // Keep this as a tiny string check so it can run on every batch item.
    private static bool ContainsNullByte(string? s) => s != null && s.IndexOf('\0') >= 0;

    internal static int SafeRun(Func<int> action, bool json = false)
    {
        if (!OfficeCli.Core.CliLogger.Enabled)
        {
            try
            {
                return action();
            }
            catch (Exception ex)
            {
                WriteError(ex, json);
                return 1;
            }
        }

        // Logging enabled: capture stdout/stderr
        var stdoutWriter = new StringWriter();
        var stderrWriter = new StringWriter();
        var origOut = Console.Out;
        var origErr = Console.Error;
        Console.SetOut(new TeeWriter(origOut, stdoutWriter));
        Console.SetError(new TeeWriter(origErr, stderrWriter));
        try
        {
            var code = action();
            var stdout = stdoutWriter.ToString().TrimEnd('\r', '\n');
            OfficeCli.Core.CliLogger.LogOutput(stdout);
            return code;
        }
        catch (Exception ex)
        {
            WriteError(ex, json);
            var stderr = stderrWriter.ToString().TrimEnd('\r', '\n');
            OfficeCli.Core.CliLogger.LogError(stderr);
            return 1;
        }
        finally
        {
            Console.SetOut(origOut);
            Console.SetError(origErr);
        }
    }

    private static void WriteError(Exception ex, bool json)
    {
        // CONSISTENCY(error-wrap): bare XmlException leaks ("Data at the root
        // level is invalid. Line 1, position 1.") when an OOXML part is
        // externally corrupted. Surface a friendlier message naming the
        // underlying cause so users know it's a malformed part, not a bug.
        var rendered = ex is System.Xml.XmlException xe
            ? new InvalidDataException(
                $"Malformed XML in document part: {xe.Message} " +
                $"(the file appears to have a corrupted OOXML part).", xe)
            : ex;
        if (json)
        {
            // JSON mode: structured error envelope to stdout so AI agents get it in the same stream
            WarningContext.End(); // discard any partial warnings
            Console.WriteLine(OutputFormatter.WrapErrorEnvelope(rendered));
        }
        else
        {
            Console.Error.WriteLine($"Error: {OfficeCli.Core.MsysPathHint.AugmentMessage(rendered.Message)}");
        }
    }

    /// <summary>
    /// Remove a path, honouring a prop-carried <c>shift</c> (Excel cell delete
    /// with shift=left|up). The CLI exposes shift via a dedicated --shift option
    /// that routes to <c>RemoveCellWithShift</c>; the MCP single-command and
    /// batch surfaces carry it inside props, so without this they silently
    /// dropped it (plain <c>Remove</c> ignores props["shift"]). Shared so all
    /// three surfaces behave identically. Returns the handler's warning (or null).
    /// </summary>
    internal static string? RemoveWithShiftSupport(OfficeCli.Core.IDocumentHandler handler, string path, Dictionary<string, string>? props)
    {
        if (props != null && props.TryGetValue("shift", out var shift) && !string.IsNullOrEmpty(shift))
        {
            if (handler is not OfficeCli.Handlers.ExcelHandler xl)
                throw new OfficeCli.Core.CliException("shift is supported only for Excel cell paths (e.g. /Sheet1/B5).")
                    { Code = "invalid_value" };
            return xl.RemoveCellWithShift(path, shift);
        }
        return handler.Remove(path, props);
    }

    /// <summary>Categorised result of <see cref="ApplySetWithCorrection"/>.</summary>
    internal sealed record SetApplyOutcome(
        List<KeyValuePair<string, string>> Applied,
        List<string> Unsupported,
        List<(string Original, string Corrected, string Value)> AutoCorrected);

    /// <summary>
    /// Apply a set's props, auto-correct any unsupported key that is a unique
    /// Levenshtein-distance-1 typo of a real prop (e.g. colot→color), and
    /// categorise the result into applied / still-unsupported / auto-corrected.
    ///
    /// This is the ONE shared core behind every set surface — the non-resident
    /// CLI set, the batch executor, the MCP single-command, and the resident —
    /// so the correction and categorisation cannot drift between them (they used
    /// to be hand-mirrored copies, flagged with `// CONSISTENCY(...)` comments).
    /// The boundary is deliberately narrow: the per-surface suggestion scope is
    /// a trivial local switch, and each caller keeps its own output envelope
    /// (CLI warnings/overlap, resident watch, batch verdict). Only the
    /// drift-prone middle is shared.
    /// </summary>
    internal static SetApplyOutcome ApplySetWithCorrection(
        OfficeCli.Core.IDocumentHandler handler, string path, Dictionary<string, string> props)
    {
        var raw = handler.Set(path, props);
        string? scope = handler switch
        {
            OfficeCli.Handlers.ExcelHandler => "excel",
            OfficeCli.Handlers.WordHandler => "word",
            OfficeCli.Handlers.PowerPointHandler => "pptx",
            _ => null,
        };
        var autoCorrected = new List<(string Original, string Corrected, string Value)>();
        var unsupported = new List<string>();
        foreach (var u in raw)
        {
            var rawKey = u.Contains(' ') ? u[..u.IndexOf(' ')] : u;
            if (props.TryGetValue(rawKey, out var val))
            {
                var (suggestion, dist, isUnique) = SuggestPropertyWithDistance(rawKey, scope);
                if (suggestion != null && dist == 1 && isUnique
                    && handler.Set(path, new Dictionary<string, string> { [suggestion] = val }).Count == 0)
                {
                    autoCorrected.Add((rawKey, suggestion, val));
                    continue;
                }
            }
            unsupported.Add(u);
        }
        // unsupported entries may carry help text ("key (valid props: ...)") or a
        // reason ("key=value (...)"); trim on the first space then split on '='
        // so the membership test matches the raw prop key.
        var unsupportedKeys = unsupported.Select(u =>
        {
            var head = u.Contains(' ') ? u[..u.IndexOf(' ')] : u;
            var eq = head.IndexOf('=');
            return eq >= 0 ? head[..eq] : head;
        }).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var autoCorrectedKeys = autoCorrected.Select(ac => ac.Original).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var applied = props.Where(kv => !unsupportedKeys.Contains(kv.Key) && !autoCorrectedKeys.Contains(kv.Key)).ToList();
        foreach (var ac in autoCorrected)
            applied.Add(new KeyValuePair<string, string>(ac.Corrected, ac.Value));
        return new SetApplyOutcome(applied, unsupported, autoCorrected);
    }

    /// <summary>
    /// Cheap single-node Format snapshot for the set-receipt normalization
    /// echo. Selector paths (no leading '/') and unresolvable paths return
    /// null — the echo is skipped rather than paying a query or guessing.
    /// </summary>
    internal static Dictionary<string, string>? TryGetFormatSnapshot(
        OfficeCli.Core.IDocumentHandler handler, string path)
    {
        if (string.IsNullOrEmpty(path) || !path.StartsWith("/")) return null;
        try
        {
            var node = handler.Get(path);
            if (node?.Format == null) return null;
            var snap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in node.Format)
                snap[kv.Key] = kv.Value switch
                {
                    null => "",
                    bool b => b ? "true" : "false",
                    _ => kv.Value.ToString() ?? "",
                };
            return snap;
        }
        catch { return null; }
    }

    /// <summary>
    /// " (applied: key=value, ...)" — the canonical form the handler actually
    /// stored, appended to the set receipt ONLY when it differs from the
    /// request (bare font → font.latin/font.ea, red → #FF0000, 14 → 14pt).
    /// A write-read-identical set keeps its receipt byte-for-byte unchanged
    /// (frozen-text discipline: extend by suffix, and only when informative).
    /// Diff entries are restricted to keys attributable to a requested key
    /// (same name or a dotted expansion of it) so recomputed unrelated
    /// Format entries can never add noise.
    /// </summary>
    internal static string BuildAppliedSuffix(
        List<KeyValuePair<string, string>> applied,
        Dictionary<string, string>? before,
        Dictionary<string, string>? after)
    {
        if (after == null || applied.Count == 0) return "";
        bool DiffAttributable(string diffKey, KeyValuePair<string, string> req) =>
            diffKey.Equals(req.Key, StringComparison.OrdinalIgnoreCase)
            || diffKey.StartsWith(req.Key + ".", StringComparison.OrdinalIgnoreCase)
            || req.Key.StartsWith(diffKey + ".", StringComparison.OrdinalIgnoreCase);
        // Fallback attribution (post-state, no diff evidence): exact key
        // match, or a dotted expansion whose value equals the SAME request's
        // value — pairing both conditions per request keeps a pre-existing
        // sibling whose value collides with a DIFFERENT request out of the
        // echo (border=thin + wrap=true must not claim
        // border.diagonalUp=true). A same-request sibling genuinely carrying
        // the requested value can still appear — value-truthful, accepted.
        bool FallbackAttributable(KeyValuePair<string, string> kv, KeyValuePair<string, string> req) =>
            kv.Key.Equals(req.Key, StringComparison.OrdinalIgnoreCase)
            || (kv.Key.StartsWith(req.Key + ".", StringComparison.OrdinalIgnoreCase)
                && string.Equals(kv.Value, req.Value, StringComparison.Ordinal));
        var diff = new List<KeyValuePair<string, string>>();
        if (before != null)
            foreach (var kv in after)
                if (!before.TryGetValue(kv.Key, out var old) || !string.Equals(old, kv.Value, StringComparison.Ordinal))
                    diff.Add(kv);
        // Resolve PER REQUEST, not per whole set: a request covered by the
        // diff uses its diff entries; a request with no diff evidence (an
        // idempotent re-write — the stored value already matched — or a first
        // write the before-snapshot couldn't resolve) falls back to the
        // post-state. Resolving the whole set at once made a key's echo
        // depend on whether some OTHER key in the same command happened to
        // change, so identical requests echoed different shapes run to run.
        var picked = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var req in applied)
        {
            var fromDiff = diff.Where(kv => DiffAttributable(kv.Key, req)).ToList();
            foreach (var kv in fromDiff.Count > 0
                ? fromDiff
                : after.Where(kv => FallbackAttributable(kv, req)))
                picked[kv.Key] = kv.Value;
        }
        var related = picked
            .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase).ToList();
        if (related.Count == 0) return "";
        // Identity — the stored form matches the request exactly: no echo.
        var identical = applied.All(req => related.Any(d =>
                d.Key.Equals(req.Key, StringComparison.OrdinalIgnoreCase)
                && string.Equals(d.Value, req.Value, StringComparison.Ordinal)))
            && related.All(d => applied.Any(req =>
                d.Key.Equals(req.Key, StringComparison.OrdinalIgnoreCase)
                && string.Equals(d.Value, req.Value, StringComparison.Ordinal)));
        if (identical) return "";
        return $" (applied: {string.Join(", ", related.Select(kv => $"{kv.Key}={kv.Value}"))})";
    }

    internal static string ExecuteBatchItem(OfficeCli.Core.IDocumentHandler handler, BatchItem item, bool json)
    {
        var format = json ? OfficeCli.Core.OutputFormat.Json : OfficeCli.Core.OutputFormat.Text;
        var props = item.Props ?? new Dictionary<string, string>();

        // Reject null bytes (U+0000) anywhere in caller-controlled strings —
        // path, selector, text, and prop values. OOXML xml writers throw
        // System.Xml.XmlException ("'.', hexadecimal value 0x00, is an
        // invalid character.") deep inside the SDK's Save path, AFTER prior
        // batch items have already mutated the document. The exception
        // bubbles up past the handler's Save and leaves the resident in a
        // state where the next close throws again — silently losing every
        // successful mutation in the same session. Reject at the boundary
        // with a stable code so the batch driver records ONE failed step
        // and keeps the rest of the document intact.
        if (ContainsNullByte(item.Path))
            throw new CliException($"path contains a NUL byte (\\u0000), which is invalid in OOXML.")
                { Code = "invalid_input" };
        if (ContainsNullByte(item.Selector))
            throw new CliException($"selector contains a NUL byte (\\u0000), which is invalid in OOXML.")
                { Code = "invalid_input" };
        if (ContainsNullByte(item.Text))
            throw new CliException($"text contains a NUL byte (\\u0000), which is invalid in OOXML.")
                { Code = "invalid_input" };
        foreach (var (pk, pv) in props)
        {
            if (ContainsNullByte(pk) || ContainsNullByte(pv))
                throw new CliException($"prop '{pk}' contains a NUL byte (\\u0000), which is invalid in OOXML.")
                    { Code = "invalid_input" };
        }

        switch (item.Command.ToLowerInvariant())
        {
            // NEWLINE-SEMANTICS-V2: version-stamp items are normally stripped
            // by BatchCompat.PrepareForReplay; tolerate one that reaches the
            // executor (plugin NDJSON lines bypass the list-level prepare).
            case "meta":
                return "meta";
            case "get":
            {
                var path = item.Path ?? "/";
                var depth = item.Depth ?? 1;
                var node = handler.Get(path, depth);
                // Error-typed nodes (e.g. namedrange not found) must surface as
                // exceptions so --stop-on-error can detect them. Without this,
                // Get returns a node with Type="error" and a message in Text,
                // ExecuteBatchItem treats it as success, and stop-on-error never fires.
                if (node.Type == "error")
                    throw new ArgumentException(node.Text ?? $"Path not found: {path}");
                // Unified envelope: batch get items emit the same
                // {matches, results: [...]} shape as query items, so callers
                // can consume batch step output with a single parser.
                if (format == OutputFormat.Json)
                    return OfficeCli.Core.OutputFormatter.FormatNodes(new List<DocumentNode> { node }, format);
                return OfficeCli.Core.OutputFormatter.FormatNode(node, format);
            }
            case "query":
            {
                // `path` is accepted as an alias for `selector` — the generic
                // field table says "path (set/remove/get target)" and users
                // carry it over to query; ignoring it silently ran an EMPTY
                // selector, i.e. returned every node as if the predicate
                // matched (the most dangerous kind of wrong data). Neither
                // field present is an error, mirroring the required CLI arg.
                var selector = item.Selector ?? item.Path ?? "";
                if (string.IsNullOrEmpty(selector))
                    throw new ArgumentException("'query' command requires 'selector' field. Example: {\"command\": \"query\", \"selector\": \"row[Score>80]\"}");
                Func<string, string>? keyResolver =
                    handler is OfficeCli.Handlers.ExcelHandler
                    && OfficeCli.Handlers.ExcelHandler.SelectorTargetsCells(selector)
                        ? OfficeCli.Handlers.ExcelHandler.ResolveCellAttributeAlias : null;
                var (results, warnings) = OfficeCli.Core.AttributeFilter.FilterSelector(selector, handler.Query, keyResolver);
                if (item.Text is { } textFilter && !string.IsNullOrEmpty(textFilter))
                    // MatchesTextFilter (not plain Contains) so a batch query
                    // text filter honours r"regex" like the CLI and resident do.
                    results = results.Where(n => n.Text != null && OfficeCli.Core.AttributeFilter.MatchesTextFilter(n.Text, textFilter)).ToList();
                foreach (var w in warnings) Console.Error.WriteLine(w.Message);
                return OfficeCli.Core.OutputFormatter.FormatNodes(results, format);
            }
            case "set":
            {
                if (string.IsNullOrEmpty(item.Path))
                    throw new ArgumentException("'set' command requires 'path' field. Example: {\"command\": \"set\", \"path\": \"/slide[1]\", \"props\": {\"bold\": \"true\"}}");
                // Match standalone `set` rejection of empty/missing props — a
                // batch step with no props is a no-op that previously reported
                // success, hiding caller mistakes (forgotten props field,
                // misspelled key promoted to root, etc.).
                if (props.Count == 0)
                    throw new ArgumentException("'set' command requires 'props' field with at least one key=value. Got empty/missing props.");
                var path = item.Path;
                OfficeCli.Core.MutationSelectorGuard.EnsureScoped(path, "set");
                // Shared core: apply + prop-autocorrect + categorise. Identical
                // across CLI / batch / MCP / resident; only the output below is
                // batch-specific.
                var (applied, unsupported, autoCorrected) = ApplySetWithCorrection(handler, path, props);
                var parts = new List<string>();
                if (autoCorrected.Count > 0)
                    parts.Add("Auto-corrected: " + string.Join(", ", autoCorrected.Select(ac => $"{ac.Original}→{ac.Corrected}")));
                if (applied.Count > 0)
                {
                    var msg = $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}";
                    if (props.ContainsKey("find"))
                    {
                        var matched = handler switch
                        {
                            OfficeCli.Handlers.WordHandler wh => wh.LastFindMatchCount,
                            OfficeCli.Handlers.PowerPointHandler ph => ph.LastFindMatchCount,
                            OfficeCli.Handlers.ExcelHandler eh => eh.LastFindMatchCount,
                            _ => 0
                        };
                        msg += $" ({matched} matched)";
                    }
                    parts.Add(msg);
                }
                if (unsupported.Count > 0)
                {
                    // /styles/<id> in Word: route through curated hints
                    // instead of the generic "use raw-set" message. raw-set
                    // is an escape hatch and pushing users there for missing
                    // curated coverage trains them out of the canonical
                    // vocabulary. See StyleUnsupportedHints.
                    if (handler is OfficeCli.Handlers.WordHandler
                        && path.StartsWith("/styles/", StringComparison.Ordinal))
                    {
                        var styleHint = OfficeCli.Core.StyleUnsupportedHints.Format(unsupported);
                        if (styleHint != null) parts.Add(styleHint);
                    }
                    else
                    {
                        string? batchScope = handler switch
                        {
                            OfficeCli.Handlers.ExcelHandler => "excel",
                            OfficeCli.Handlers.WordHandler => "word",
                            OfficeCli.Handlers.PowerPointHandler => "pptx",
                            _ => null,
                        };
                        parts.Add(FormatUnsupported(unsupported, batchScope));
                    }
                    // Mirror standalone `set`'s allFailed semantics: if every
                    // requested property was rejected (nothing applied), the
                    // step is a failure, not a successful no-op. Without
                    // this, batch swallowed unsupported_property into an inner
                    // success=true while the same set issued via the
                    // standalone set command returned success=false exit 2.
                    // Per-step verdict flips to false; outer batch envelope
                    // still rides on the existing partial-success rule.
                    if (applied.Count == 0)
                        throw new CliException(string.Join("\n", parts)) { Code = "unsupported_property" };
                }
                return string.Join("\n", parts);
            }
            case "add":
            {
                var parentPath = item.Parent ?? item.Path;
                if (string.IsNullOrEmpty(parentPath))
                    throw new ArgumentException("'add' command requires 'parent' field. Example: {\"command\": \"add\", \"parent\": \"/slide[1]\", \"type\": \"shape\", \"props\": {\"text\": \"Hello\"}}");
                if (string.IsNullOrEmpty(item.Type) && string.IsNullOrEmpty(item.From))
                    throw new ArgumentException("'add' command requires 'type' or 'from' field. Example: {\"command\": \"add\", \"parent\": \"/\", \"type\": \"slide\"}");
                InsertPosition? pos = null;
                if (item.Index.HasValue) pos = InsertPosition.AtIndex(item.Index.Value);
                else if (!string.IsNullOrEmpty(item.After)) pos = InsertPosition.AfterElement(item.After);
                else if (!string.IsNullOrEmpty(item.Before)) pos = InsertPosition.BeforeElement(item.Before);

                if (!string.IsNullOrEmpty(item.From))
                {
                    var resultPath = handler.CopyFrom(item.From, parentPath, pos);
                    return $"Copied to {resultPath}";
                }
                else
                {
                    var type = item.Type ?? "";
                    // Wrap props in a tracking dict (matches CLI/resident add): a
                    // key the handler reads is consumed, so UnusedKeys after Add
                    // is the generic unsupported-prop set across ALL handlers.
                    // Previously batch/MCP add saw only Word's curated
                    // LastAddUnsupportedProps, silently dropping an unknown prop
                    // on a pptx/xlsx add that the CLI/resident would report.
                    var tracking = new OfficeCli.Core.TrackingPropertyDictionary(props);
                    var resultPath = handler.Add(parentPath, type, pos, tracking);
                    var addMsg = $"Added {type} at {resultPath}";
                    var addUnsupported = tracking.UnusedKeys.ToList();
                    if (handler is OfficeCli.Handlers.WordHandler addWh)
                        addUnsupported.AddRange(addWh.LastAddUnsupportedProps);
                    if (addUnsupported.Count > 0)
                    {
                        // Word → curated hints (keyed off the result path so a
                        // /styles add gets style vocabulary); other handlers →
                        // the generic scoped formatter.
                        string? hint;
                        if (handler is OfficeCli.Handlers.WordHandler)
                            hint = OfficeCli.Core.StyleUnsupportedHints.Format(addUnsupported, ScopeLabelForWordPath(resultPath));
                        else
                        {
                            string? addScope = handler is OfficeCli.Handlers.ExcelHandler ? "excel"
                                : handler is OfficeCli.Handlers.PowerPointHandler ? "pptx" : null;
                            hint = FormatUnsupported(addUnsupported, addScope);
                        }
                        if (hint != null) addMsg += "\nWARNING: " + hint;
                    }
                    return addMsg;
                }
            }
            case "import":
            {
                // CSV/TSV bulk import — batch counterpart of the standalone
                // `officecli import` command (CommandBuilder.Import.cs). The
                // CSV content rides the item's `text` field; `parent` is the
                // sheet path. This is the value-baseline carrier for
                // `dump --format batch` on .xlsx (ExcelBatchEmitter).
                if (handler is not OfficeCli.Handlers.ExcelHandler importXl)
                    throw new CliException("'import' batch command is only supported for .xlsx files")
                        { Code = "unsupported_type" };
                var importParent = item.Parent ?? item.Path;
                if (string.IsNullOrEmpty(importParent))
                    throw new ArgumentException("'import' command requires 'parent' field (sheet path). Example: {\"command\": \"import\", \"parent\": \"/Sheet1\", \"text\": \"a,b\\n1,2\"}");
                if (item.Text == null)
                    throw new ArgumentException("'import' command requires 'text' field with the CSV/TSV content.");
                // CONSISTENCY(import-vocabulary): props mirror the standalone
                // command's options — format=csv|tsv, header, start-cell.
                char importDelim = ',';
                if (props.TryGetValue("format", out var importFmt) && !string.IsNullOrEmpty(importFmt))
                {
                    importDelim = importFmt.ToLowerInvariant() switch
                    {
                        "tsv" => '\t',
                        "csv" => ',',
                        _ => throw new CliException($"Unknown format: {importFmt}. Use 'csv' or 'tsv'")
                            { Code = "invalid_value", ValidValues = ["csv", "tsv"] },
                    };
                }
                var importHeader = props.TryGetValue("header", out var importHdr)
                    && OfficeCli.Core.ParseHelpers.IsTruthy(importHdr);
                var importStart = props.TryGetValue("start-cell", out var importSc) && !string.IsNullOrEmpty(importSc)
                    ? importSc
                    : props.TryGetValue("startcell", out var importSc2) && !string.IsNullOrEmpty(importSc2)
                        ? importSc2 : "A1";
                return importXl.Import(importParent, item.Text, importDelim, importHeader, importStart);
            }
            case "remove":
            {
                if (string.IsNullOrEmpty(item.Path))
                    throw new ArgumentException("'remove' command requires 'path' field. Example: {\"command\": \"remove\", \"path\": \"/slide[1]/shape[2]\"}");
                var path = item.Path;
                OfficeCli.Core.MutationSelectorGuard.EnsureScoped(path, "remove");
                var warning = RemoveWithShiftSupport(handler, path, item.Props);
                var msg = $"Removed {path}";
                if (warning != null) msg += $"\n{warning}";
                return msg;
            }
            case "move":
            {
                var path = item.Path ?? "/";
                InsertPosition? movePos = null;
                if (item.Index.HasValue) movePos = InsertPosition.AtIndex(item.Index.Value);
                else if (!string.IsNullOrEmpty(item.After)) movePos = InsertPosition.AfterElement(item.After);
                else if (!string.IsNullOrEmpty(item.Before)) movePos = InsertPosition.BeforeElement(item.Before);
                // Pass props to the 4-arg Move like the CLI and resident do; the
                // batch/MCP path previously dropped move-time properties.
                var resultPath = handler.Move(path, item.To, movePos, props.Count > 0 ? props : null);
                return $"Moved to {resultPath}";
            }
            case "swap":
            {
                // Second element: accept `path2` (canonical — the single-command
                // MCP tool and the CLI `swap path1 path2` both use it) or the
                // legacy `to`. Before path2 was carried, an agent that learned
                // swap from the single command produced a batch item that
                // silently failed the path-presence check below.
                var swapTo = !string.IsNullOrEmpty(item.Path2) ? item.Path2 : item.To;
                if (string.IsNullOrEmpty(item.Path) || string.IsNullOrEmpty(swapTo))
                    throw new ArgumentException("'swap' command requires 'path' and 'path2' (or 'to') fields. Example: {\"command\": \"swap\", \"path\": \"/slide[1]\", \"path2\": \"/slide[2]\"}");
                var (p1, p2) = handler switch
                {
                    OfficeCli.Handlers.PowerPointHandler ppt => ppt.Swap(item.Path, swapTo),
                    OfficeCli.Handlers.WordHandler word => word.Swap(item.Path, swapTo),
                    OfficeCli.Handlers.ExcelHandler excel => excel.Swap(item.Path, swapTo),
                    _ => throw new InvalidOperationException("swap not supported for this document type")
                };
                return $"Swapped {p1} <-> {p2}";
            }
            case "view":
            {
                var mode = item.Mode ?? "text";
                if (mode.ToLowerInvariant() is "html" or "h")
                {
                    if (handler is OfficeCli.Handlers.PowerPointHandler pptH)
                        return pptH.ViewAsHtml();
                    if (handler is OfficeCli.Handlers.ExcelHandler excelH)
                        return excelH.ViewAsHtml();
                    if (handler is OfficeCli.Handlers.WordHandler wordH)
                        return wordH.ViewAsHtml();
                }
                if (mode.ToLowerInvariant() is "svg" or "g" && handler is OfficeCli.Handlers.PowerPointHandler pptSvg)
                {
                    return pptSvg.ViewAsSvg(1);
                }
                return mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(null, null, null, null),
                    "annotated" or "a" => handler.ViewAsAnnotated(null, null, null, null),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OfficeCli.Core.OutputFormatter.FormatIssues(handler.ViewAsIssues(null, null), format),
                    _ => $"Unknown mode: {mode}"
                };
            }
            case "raw":
            {
                if (string.IsNullOrEmpty(item.Part))
                    throw new ArgumentException("'raw' command requires 'part' field. Example: {\"command\": \"raw\", \"part\": \"/document\"} (docx), {\"command\": \"raw\", \"part\": \"/presentation\"} (pptx), {\"command\": \"raw\", \"part\": \"/sheet[1]\"} (xlsx)");
                return handler.Raw(item.Part, null, null, null);
            }
            case "raw-set":
            {
                var partPath = item.Part ?? "/document";
                var xpath = item.Xpath ?? "";
                var action = item.Action ?? "";
                handler.RawSet(partPath, xpath, action, item.Xml);
                return $"raw-set {action} applied";
            }
            case "add-part":
            {
                if (string.IsNullOrEmpty(item.Parent))
                    throw new ArgumentException("'add-part' command requires 'parent' field. Example: {\"command\": \"add-part\", \"parent\": \"/slide[1]\", \"type\": \"smartart\", \"props\": {\"data\": \"rId2\"}}");
                if (string.IsNullOrEmpty(item.Type))
                    throw new ArgumentException("'add-part' command requires 'type' field. Supported (pptx): chart, smartart, video, audio, model3d, ole, image, hyperlink, theme.");
                var (relId, partOut) = handler.AddPart(item.Parent, item.Type, props);
                return $"Created {item.Type} part: relId={relId} path={partOut}";
            }
            case "validate":
            {
                var errors = handler.Validate();
                if (errors.Count == 0) return "Validation passed: no errors found.";
                var lines = new List<string> { $"Found {errors.Count} validation error(s):" };
                foreach (var err in errors)
                {
                    lines.Add($"  [{err.ErrorType}] {err.Description}");
                    if (err.Path != null) lines.Add($"    Path: {err.Path}");
                    if (err.Part != null) lines.Add($"    Part: {err.Part}");
                }
                return string.Join("\n", lines);
            }
            default:
                if (string.IsNullOrEmpty(item.Command))
                    throw new InvalidOperationException(
                        "Batch item missing required 'command' field. " +
                        "Valid commands: get, query, set, add, remove, move, view, raw, validate. " +
                        "Example: {\"command\": \"set\", \"path\": \"/Sheet1/A1\", \"props\": {\"value\": \"hello\"}}");
                // A "command" containing whitespace is almost always a whole CLI
                // line stuffed into the verb field (e.g. "add /slide[1] --type
                // shape --prop ...") — the single most common batch-item mistake.
                // Diagnose it specifically and point at the item schema; a plain
                // unknown verb just gets the schema pointer.
                var batchHint = item.Command.Any(char.IsWhiteSpace)
                    ? " — that looks like a whole CLI line placed in \"command\". Use the bare verb only and put the"
                      + " rest in sibling fields, e.g. {\"command\":\"add\",\"parent\":\"/slide[1]\",\"type\":\"shape\","
                      + "\"props\":{...}}. Run `help batch` for the item schema."
                    : " Run `help batch` for the JSON item schema.";
                throw new InvalidOperationException($"Unknown command: '{item.Command}'. Valid commands: get, query, set, add, remove, move, swap, view, raw, validate.{batchHint}");
        }
    }

    private static Dictionary<string, string> ParsePropsArray(string[]? props)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in props ?? Array.Empty<string>())
        {
            var eqIdx = prop.IndexOf('=');
            // BUG-R40-B12: previously `eqIdx > 0` silently dropped both
            // `--prop =value` (empty key, eqIdx==0) and `--prop key`
            // (no equals, eqIdx==-1). Surface the empty-key form as a
            // hard error so AI callers don't waste a turn wondering why
            // their property had no effect.
            if (eqIdx == 0)
                throw new ArgumentException(
                    $"Invalid --prop '{prop}': key is empty. Use key=value (e.g. --prop name=Title).");
            if (eqIdx > 0)
            {
                var key = prop[..eqIdx];
                var value = prop[(eqIdx + 1)..];
                // CONSISTENCY(text-escape-boundary): C-style escape resolution
                // (\\n, \\t, \\r, \\\\) is a CLI-input concern only. The shell
                // gives us the literal four-character sequence `\\n` which a
                // user typing `--prop text='line1\\nline2'` plainly wants as
                // a newline. Handlers no longer call TextEscape.Resolve
                // internally — that double-resolution mangled batch JSON
                // payloads, where `"text": "hello\\nworld"` already arrives
                // as `hello\\nworld` literal after JSON parsing and must NOT
                // be turned into a newline. Affected keys are the text-valued
                // props: `text`, `value`, and the row-level `c1…cN` cell-text
                // shortcuts (so `--prop c1='a\nb'` breaks the line exactly like
                // `--prop text=` does); other props (colors, paths, numbers)
                // are passed through untouched.
                if (KeyTakesCEscapes(key))
                {
                    value = OfficeCli.Core.TextEscape.Resolve(value);
                }
                dict[key] = value;
            }
        }

        // NEWLINE-SEMANTICS-V2 + CONSISTENCY(text-escape-boundary): find /
        // replace get the same C-escape convenience as text= (`--find '\v'
        // --replace '\n'` works without shell $'..' quoting) — EXCEPT when
        // the find is a regex (r"..." prefix or regex=true): regex has its
        // own escape language (\\., \b, \d) that C-escape resolution would
        // corrupt, so regex invocations are passed through verbatim, same
        // stance as Word's wildcard mode. Post-pass here (not in the per-key
        // loop) because the decision needs the regex key's value.
        bool findIsRegex =
            (dict.TryGetValue("regex", out var rxFlag) && OfficeCli.Core.ParseHelpers.IsTruthySafe(rxFlag))
            || (dict.TryGetValue("find", out var fv)
                && (fv.StartsWith("r\"", StringComparison.Ordinal) || fv.StartsWith("r'", StringComparison.Ordinal)));
        if (!findIsRegex)
        {
            if (dict.TryGetValue("find", out var findVal))
                dict["find"] = OfficeCli.Core.TextEscape.Resolve(findVal);
            if (dict.TryGetValue("replace", out var replVal))
                dict["replace"] = OfficeCli.Core.TextEscape.Resolve(replVal);
        }
        return dict;
    }

    /// <summary>
    /// CONSISTENCY(text-escape-boundary): the --prop keys whose values go
    /// through TextEscape.Resolve on the way in. Anything that builds argv from
    /// a payload where escapes are ALREADY literal — a JSON body, a batch item —
    /// must run those same values through TextEscape.Protect first, or the
    /// resolution happens a second time and eats the user's backslashes.
    /// </summary>
    internal static bool KeyTakesCEscapes(string key)
        => key.Equals("text", StringComparison.OrdinalIgnoreCase)
           || key.Equals("value", StringComparison.OrdinalIgnoreCase)
           || IsCellTextShortcutKey(key);

    // Row-level cell-text shortcut key: `c` followed by digits (c1, c2, …, cN).
    // These carry table-cell text, so they take the same `\n`/`\t` escape
    // resolution as `text=` (see CONSISTENCY(text-escape-boundary) above).
    private static bool IsCellTextShortcutKey(string key)
    {
        if (key.Length < 2 || (key[0] != 'c' && key[0] != 'C')) return false;
        for (int i = 1; i < key.Length; i++)
            if (!char.IsDigit(key[i])) return false;
        return true;
    }

    internal static void PrintBatchResults(List<BatchResult> results, bool json, int totalCount = 0, TextWriter? output = null, bool atomicRolledBack = false)
    {
        var @out = output ?? Console.Out;
        if (totalCount == 0) totalCount = results.Count;

        if (json)
        {
            var succeeded = results.Count(r => r.Success);
            var failed = results.Count - succeeded;
            var skipped = totalCount - results.Count;

            using var ms = new System.IO.MemoryStream();
            using (var writer = new System.Text.Json.Utf8JsonWriter(ms))
            {
                writer.WriteStartObject();
                writer.WritePropertyName("results");
                System.Text.Json.JsonSerializer.Serialize(writer, results, BatchJsonContext.Default.ListBatchResult);
                writer.WriteStartObject("summary");
                writer.WriteNumber("total", totalCount);
                writer.WriteNumber("executed", results.Count);
                writer.WriteNumber("succeeded", succeeded);
                writer.WriteNumber("failed", failed);
                writer.WriteNumber("skipped", skipped);
                // Additive field, only present when the atomic default
                // discarded the batch — parsers keying on the existing
                // summary fields are unaffected.
                if (atomicRolledBack) writer.WriteBoolean("atomicRolledBack", true);
                writer.WriteEndObject();
                writer.WriteEndObject();
            }

            var fullBytes = ms.ToArray();
            if (fullBytes.Length <= 8192)
            {
                @out.WriteLine(System.Text.Encoding.UTF8.GetString(fullBytes));
            }
            else
            {
                // Spill full output to temp file
                var tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officecli_batch_{Guid.NewGuid():N}.json");
                System.IO.File.WriteAllBytes(tempPath, fullBytes);

                // Write slim envelope
                using var slimMs = new System.IO.MemoryStream();
                using (var slimWriter = new System.Text.Json.Utf8JsonWriter(slimMs))
                {
                    slimWriter.WriteStartObject();
                    slimWriter.WriteString("outputFile", tempPath);
                    slimWriter.WriteNumber("outputSize", fullBytes.Length);
                    slimWriter.WriteStartArray("results");
                    foreach (var r in results)
                    {
                        slimWriter.WriteStartObject();
                        slimWriter.WriteNumber("index", r.Index);
                        slimWriter.WriteBoolean("success", r.Success);
                        if (r.Error != null)
                        {
                            slimWriter.WriteString("error", r.Error);
                            if (r.Code != null)
                                slimWriter.WriteString("code", r.Code);
                            if (r.Item != null)
                            {
                                slimWriter.WritePropertyName("item");
                                System.Text.Json.JsonSerializer.Serialize(slimWriter, r.Item, BatchJsonContext.Default.BatchItem);
                            }
                        }
                        slimWriter.WriteEndObject();
                    }
                    slimWriter.WriteEndArray();
                    slimWriter.WriteStartObject("summary");
                    slimWriter.WriteNumber("total", totalCount);
                    slimWriter.WriteNumber("executed", results.Count);
                    slimWriter.WriteNumber("succeeded", succeeded);
                    slimWriter.WriteNumber("failed", failed);
                    slimWriter.WriteNumber("skipped", skipped);
                    if (atomicRolledBack) slimWriter.WriteBoolean("atomicRolledBack", true);
                    slimWriter.WriteEndObject();
                    slimWriter.WriteEndObject();
                }
                @out.WriteLine(System.Text.Encoding.UTF8.GetString(slimMs.ToArray()));
            }
        }
        else
        {
            for (int i = 0; i < results.Count; i++)
            {
                var r = results[i];
                var prefix = $"[{i + 1}] ";
                if (r.Success)
                {
                    if (!string.IsNullOrEmpty(r.Output))
                        @out.WriteLine($"{prefix}{r.Output}");
                    else
                        @out.WriteLine($"{prefix}OK");
                }
                else
                {
                    @out.WriteLine($"{prefix}ERROR: {r.Error}");
                }
            }

            var succeeded = results.Count(r => r.Success);
            var failed = results.Count - succeeded;
            // FROZEN TEXT: the "Batch complete: N succeeded, M failed" skeleton
            // is a machine-consumed contract — extend by SUFFIX only.
            var atomicNote = atomicRolledBack ? " (atomic: no changes were applied)" : "";
            @out.WriteLine($"\nBatch complete: {succeeded} succeeded, {failed} failed, {results.Count} total{atomicNote}");
        }
    }

    private static string FormatValidationErrors(List<ValidationError> errors)
    {
        var sb = new StringBuilder();
        sb.Append("{\"count\":").Append(errors.Count).Append(",\"errors\":[");
        for (int i = 0; i < errors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var e = errors[i];
            sb.Append("{\"type\":\"").Append(EscapeJson(e.ErrorType)).Append('"');
            sb.Append(",\"description\":\"").Append(EscapeJson(e.Description)).Append('"');
            if (e.Path != null) sb.Append(",\"path\":\"").Append(EscapeJson(e.Path)).Append('"');
            if (e.Part != null) sb.Append(",\"part\":\"").Append(EscapeJson(e.Part)).Append('"');
            sb.Append('}');
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static string EscapeJson(string s) => s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");

    internal static List<CliWarning>? ReportNewErrorsAsWarnings(OfficeCli.Core.IDocumentHandler handler, HashSet<string> errorsBefore)
    {
        var errorsAfter = handler.Validate();
        var newErrors = errorsAfter.Where(e => !errorsBefore.Contains(e.Description)).ToList();
        if (newErrors.Count == 0) return null;
        return newErrors.Select(err => new CliWarning
        {
            Message = $"[{err.ErrorType}] {err.Description}" +
                (err.Path != null ? $" (Path: {err.Path})" : "") +
                (err.Part != null ? $" (Part: {err.Part})" : ""),
            Code = "validation_error"
        }).ToList();
    }

    internal static void ReportNewErrors(OfficeCli.Core.IDocumentHandler handler, HashSet<string> errorsBefore, List<CliWarning>? preComputed = null)
    {
        var warnings = preComputed ?? ReportNewErrorsAsWarnings(handler, errorsBefore);
        if (warnings is { Count: > 0 })
        {
            Console.WriteLine($"VALIDATION: {warnings.Count} new error(s) introduced:");
            foreach (var w in warnings)
                Console.WriteLine($"  {w.Message}");
        }
    }

    /// <summary>
    /// Detect bare key=value tokens and --key value flag patterns in unmatched arguments (user forgot --prop).
    /// Returns a list of "key=value" strings suitable for --prop suggestions.
    /// </summary>
    internal static List<string> DetectUnmatchedKeyValues(System.CommandLine.ParseResult parseResult)
    {
        var result = new List<string>();
        var tokens = parseResult.UnmatchedTokens;
        var knownPropsLower = new HashSet<string>(KnownProps.Select(p => p.ToLowerInvariant()));

        for (int i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];

            // Pattern 1: bare key=value (e.g. "text=Hello")
            if (System.Text.RegularExpressions.Regex.IsMatch(token, @"^[A-Za-z_.][A-Za-z0-9_.]*=.+$"))
            {
                result.Add(token);
                continue;
            }

            // Pattern 2: --key value (e.g. "--text Hello" or "--fill yellow")
            // Only match if the key (without --) is a known property name
            if (token.StartsWith("--") && token.Length > 2)
            {
                var key = token[2..];
                if (knownPropsLower.Contains(key.ToLowerInvariant()) && i + 1 < tokens.Count)
                {
                    var nextToken = tokens[i + 1];
                    // Don't consume the next token if it also looks like a flag
                    if (!nextToken.StartsWith("--"))
                    {
                        result.Add($"{key}={nextToken}");
                        i++; // skip the value token
                        continue;
                    }
                }
            }

            // Pattern 3 (BUG-BT-R6): common typos for the `--prop` option name.
            // `--props '{"k":"v"}'` is silently swallowed by System.CommandLine
            // because `--props` (with trailing s) is not a known option, so the
            // JSON value goes into UnmatchedTokens too. Catch the typo so the
            // existing warning machinery emits a clear hint instead of letting
            // the agent ship a shape with no text.
            if (token is "--props" or "-props" or "--prop=" && i + 1 < tokens.Count)
            {
                var nextToken = tokens[i + 1];
                if (!nextToken.StartsWith("--"))
                {
                    result.Add($"--prop {nextToken}");
                    i++;
                    continue;
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Hard-reject any unmatched `--option` token that
    /// <see cref="DetectUnmatchedKeyValues"/> did not convert into a
    /// missing-prop warning. Commands parsed with
    /// TreatUnmatchedTokensAsErrors=false otherwise swallow unknown flags
    /// (e.g. `add ... --at A2`) silently with exit 0 — the element lands
    /// somewhere the caller did not intend and nothing surfaces the typo.
    /// </summary>
    internal static void RejectUnknownOptionTokens(
        System.CommandLine.ParseResult parseResult, List<string> claimedKeyValues)
    {
        var tokens = parseResult.UnmatchedTokens;
        var claimedKeys = new HashSet<string>(
            claimedKeyValues.Select(kv => kv.Split('=', 2)[0].Trim().TrimStart('-')),
            StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token == "--") break;                       // explicit passthrough separator
            if (!token.StartsWith("--") || token.Length <= 2) continue;
            var key = token[2..];
            if (key.Contains('='))                          // --key=value form
                key = key[..key.IndexOf('=')];
            if (claimedKeys.Contains(key)) continue;        // already warned as missing --prop
            if (key is "props" or "prop") continue;         // typo forms handled above
            var valueHint = i + 1 < tokens.Count && !tokens[i + 1].StartsWith("--")
                ? $"{key}={tokens[i + 1]}"
                : $"{key}=<value>";
            throw new OfficeCli.Core.CliException($"Unrecognized option '{token}'.")
            {
                Code = "invalid_argument",
                Suggestion = $"Element properties are passed via --prop, e.g. --prop {valueHint}. Run 'officecli add --help' for the supported options."
            };
        }
    }

    /// <summary>
    /// Reduce a Word handler result path to the meaningful scope label for
    /// UNSUPPORTED messages — "/styles", "/body/p[N]", "/body/p[N]/r[N]".
    /// Stops at the first segment that is not a known top-level Word
    /// container so unfamiliar paths fall back to the full path.
    /// </summary>
    private static string ScopeLabelForWordPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return "/";
        if (path.StartsWith("/styles/", StringComparison.Ordinal)) return "/styles";
        // Trim everything past the last bracketed-segment we recognize for
        // paragraph/run paths. Keep the path as-is for everything else.
        return path;
    }

    internal static string FormatUnsupported(IEnumerable<string> unsupported, string? scope = null)
    {
        var parts = new List<string>();
        foreach (var prop in unsupported)
        {
            // Word scoped-alternative hints (e.g. cantSplit rejected at table
            // level → "row-scoped: …"). Mirrors the resident add path, which
            // routes through StyleUnsupportedHints.Format directly.
            if (scope == "word" && !prop.Contains('(')
                && OfficeCli.Core.StyleUnsupportedHints.TryGetHint(prop, out var scopedHint))
            {
                parts.Add($"{prop} ({scopedHint})");
                continue;
            }
            // An entry that already carries a handler-embedded hint
            // ("fillBg (background of ...)", "cap (valid cell props: ...)")
            // doesn't need a did-you-mean guess stacked on top.
            if (prop.Contains('('))
            {
                parts.Add(prop);
                continue;
            }
            var suggestion = SuggestPropertyScoped(prop, scope);
            parts.Add(suggestion != null ? $"{prop} (did you mean: {suggestion}?)" : prop);
        }
        return $"UNSUPPORTED props: {string.Join(", ", parts)}. Run 'officecli help <format> <element>' to see valid props, or raw-set for raw XML.";
    }

    /// <summary>
    /// Property keys that belong to PPTX shape/text semantics and should not
    /// be offered as suggestions when the caller is operating on an Excel
    /// document (R2-4). Keep the list conservative — only keys whose presence
    /// in an Excel error message would be clearly misleading.
    /// </summary>
    internal static readonly HashSet<string> PptxOnlyProps = new(StringComparer.OrdinalIgnoreCase)
    {
        "rotation", "opacity", "glow", "shadow",
        "firstSliceAngle", "holeSize", "bubbleScale", "explosion",
        "view3d", "varyColors",
    };

    /// <summary>
    /// Property keys exclusive to Word document-level concerns that should
    /// not bleed into Excel suggestions.
    /// </summary>
    internal static readonly HashSet<string> WordOnlyProps = new(StringComparer.OrdinalIgnoreCase)
    {
        "pageWidth", "pageHeight", "orientation",
    };

    internal static readonly string[] KnownProps = new[]
    {
        "text", "bold", "italic", "underline", "strike", "font", "size", "color",
        "highlight", "alignment", "spacing", "indent", "shd", "border",
        "width", "height", "valign", "header", "formula", "value", "type",
        "fill", "src", "path", "title", "name", "style", "caps", "smallcaps",
        "lineSpacing", "listStyle", "start", "level", "cols", "rows",
        "gridspan", "vmerge", "nowrap", "padding", "margin",
        "orientation", "pageWidth", "pageHeight",
        "x", "y", "cx", "cy", "rotation", "opacity",
        "border.color", "border.width", "border.style",
        "font.color", "font.size", "font.name", "font.bold", "font.italic",
        "hyperlink", "link", "tooltip", "alt", "description",
        "font.strike", "font.underline", "tabColor", "shadow", "glow", "numberformat",
        // Chart properties
        "chartType", "title", "legend", "dataLabels", "labelPos", "labelFont",
        "axisFont", "axisTitle", "catTitle", "axisMin", "axisMax", "majorUnit", "minorUnit",
        "axisNumFmt", "axisVisible", "majorTickMark", "minorTickMark", "tickLabelPos",
        "axisPosition", "crosses", "crossesAt", "crossBetween", "axisOrientation", "logBase",
        "dispUnits", "labelOffset", "tickLabelSkip",
        "gridlines", "minorGridlines", "plotFill", "chartFill",
        "colors", "gradient", "gradients", "lineWidth", "lineDash",
        "marker", "markerSize", "transparency", "smooth", "showMarker",
        "scatterStyle", "radarStyle", "varyColors", "dispBlanksAs",
        "roundedCorners", "plotVisOnly", "trendline", "invertIfNeg", "explosion",
        "errBars", "gapWidth", "overlap", "secondaryAxis", "dataTable",
        "firstSliceAngle", "holeSize", "bubbleScale", "shape", "gapDepth",
        "dropLines", "hiLowLines", "upDownBars", "serLines",
        "plotArea.border", "chartArea.border", "legend.overlay",
        "plotArea.x", "plotArea.y", "plotArea.w", "plotArea.h",
        "title.x", "title.y", "title.w", "title.h",
        "legend.x", "legend.y", "legend.w", "legend.h",
        "datalabels.separator", "datalabels.numfmt", "leaderLines",
        "view3d", "categories", "data",
        "referenceLine", "refLine", "targetLine", "preset", "colorRule",
        "conditionalColor", "comboTypes", "axisLine",
    };

    internal static string? SuggestProperty(string input)
    {
        var (best, _, _) = SuggestPropertyWithDistance(input);
        return best;
    }

    /// <summary>
    /// Scoped variant: filters the suggestion pool against a target document
    /// format ("excel", "word", "pptx", or null for unscoped) to avoid
    /// cross-format leakage such as suggesting PPTX 'rotation' for an
    /// Excel pivot property (R2-4).
    /// </summary>
    internal static string? SuggestPropertyScoped(string input, string? scope)
    {
        var (best, _, _) = SuggestPropertyWithDistance(input, scope);
        return best;
    }

    /// <summary>
    /// Returns (bestMatch, distance, isUnique) where isUnique means no other candidate shares the same distance.
    /// </summary>
    internal static (string? Best, int Distance, bool IsUnique) SuggestPropertyWithDistance(string input, string? scope = null)
    {
        // Strip help text suffix if present (e.g. "key (valid props: ...)")
        var rawInput = input.Contains(' ') ? input[..input.IndexOf(' ')] : input;
        var lower = rawInput.ToLowerInvariant();

        // Table cell-content keys are 1-based (r1c1, r1c2, …) across all
        // handlers (pptx AddTable, word AddTable). A 0-based r0c0 / cN starting
        // at 0 is the single most common miss. Point straight
        // at the 1-based form rather than letting Levenshtein guess a far-off
        // KnownProp.
        var rcMatch = System.Text.RegularExpressions.Regex.Match(lower, @"^r(\d+)c(\d+)$");
        if (rcMatch.Success)
        {
            var rr = int.Parse(rcMatch.Groups[1].Value);
            var cc = int.Parse(rcMatch.Groups[2].Value);
            if (rr == 0 || cc == 0)
                return ($"r{rr + 1}c{cc + 1} (cell keys are 1-based)", 1, true);
        }
        string? best = null;
        int bestDist = int.MaxValue;
        int bestCount = 0; // how many props share the best distance

        HashSet<string>? exclude = null;
        switch (scope?.ToLowerInvariant())
        {
            case "excel":
                exclude = new HashSet<string>(PptxOnlyProps, StringComparer.OrdinalIgnoreCase);
                foreach (var w in WordOnlyProps) exclude.Add(w);
                break;
            case "word":
                exclude = PptxOnlyProps;
                break;
            case "pptx":
                exclude = WordOnlyProps;
                break;
        }

        foreach (var prop in KnownProps)
        {
            if (exclude != null && exclude.Contains(prop)) continue;
            var dist = OfficeCli.Core.EditDistance.Damerau(lower, prop.ToLowerInvariant());
            if (dist > 0 && dist <= Math.Max(2, rawInput.Length / 3))
            {
                if (dist < bestDist)
                {
                    bestDist = dist;
                    best = prop;
                    bestCount = 1;
                }
                else if (dist == bestDist)
                {
                    bestCount++;
                }
            }
        }

        return best != null ? (best, bestDist, bestCount == 1) : (null, int.MaxValue, false);
    }

    // ==================== PPT spatial info helpers ====================

    /// <summary>
    /// Check if a .docx file has document protection enforced.
    /// Returns 0 if no protection or if the path targets an editable element.
    /// Returns 1 with error output if the document is protected and the target is not an editable region.
    /// </summary>
    private static int CheckDocxProtection(string filePath, string path, bool json)
    {
        try
        {
            using var handler = DocumentHandlerFactory.Open(filePath, editable: false);
            var root = handler.Get("/");
            var protection = root.Format.TryGetValue("protection", out var pVal) ? pVal?.ToString() : "none";
            var enforced = root.Format.TryGetValue("protectionEnforced", out var eVal) && eVal is true;

            if (!enforced || protection == "none")
                return 0;

            // Allow writes to formfield and SDT paths (they handle their own editable check)
            if (path.StartsWith("/formfield[", StringComparison.OrdinalIgnoreCase))
                return 0;
            if (path.Contains("/sdt[", StringComparison.OrdinalIgnoreCase))
                return 0;

            // Document is protected — block the write
            var msg = $"Document is protected (mode: {protection}). " +
                      "Use Query(\"editable\") to find editable fields, or use --force to override protection.";
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg, new List<OfficeCli.Core.CliWarning>()));
            else
                Console.Error.WriteLine($"ERROR: {msg}");
            return 1;
        }
        catch
        {
            // If we can't read protection info, allow the write to proceed
            return 0;
        }
    }

    // Batch-scoped protection gate evaluated against an ALREADY-OPEN handler's
    // in-memory DOM — no extra file open, and authoritative even when the
    // on-disk copy lags the in-memory tree (resident sessions flush only on
    // save/close, so a prior in-memory protection change would not yet be on
    // disk; reading the file there could mis-gate the batch). This is the
    // in-memory equivalent of CheckDocxProtection applied once per batch: a
    // protected document rejects the batch unless every gated mutation targets
    // a formfield/sdt path (which manage their own editable check). Returns 0
    // to allow, non-zero after emitting the rejection.
    internal static string? GetBatchProtectionBlock(OfficeCli.Core.IDocumentHandler handler, List<BatchItem> items)
    {
        OfficeCli.Core.DocumentNode root;
        try { root = handler.Get("/"); }
        catch { return null; } // can't read protection -> allow (mirrors CheckDocxProtection)
        var protection = root.Format.TryGetValue("protection", out var pVal) ? pVal?.ToString() : "none";
        var enforced = root.Format.TryGetValue("protectionEnforced", out var eVal) && eVal is true;
        if (!enforced || protection == "none")
            return null;
        foreach (var item in items)
        {
            var cmd = (item.Command ?? "").ToLowerInvariant();
            if (cmd is not ("set" or "add" or "remove" or "raw-set"))
                continue;
            if (item.Props != null && item.Props.Keys.Any(k =>
                k.Equals("protection", StringComparison.OrdinalIgnoreCase)))
                continue;
            var path = item.Path ?? "";
            if (path.StartsWith("/formfield[", StringComparison.OrdinalIgnoreCase))
                continue;
            if (path.Contains("/sdt[", StringComparison.OrdinalIgnoreCase))
                continue;
            // Non-exempt mutation against a protected document — block the batch.
            return $"Document is protected (mode: {protection}). " +
                   "Use Query(\"editable\") to find editable fields, or use --force to override protection.";
        }
        return null;
    }

    private static readonly HashSet<string> PositionKeys = new(StringComparer.OrdinalIgnoreCase)
        { "x", "left", "y", "top", "width", "w", "height", "h" };

    /// <summary>
    /// For PPT spatial elements, return coordinate string like "x: 0cm  y: 5cm  width: 33.87cm  height: 5cm".
    /// Returns null for non-spatial elements (slide, Word, Excel).
    /// </summary>
    private static string? GetPptSpatialLine(IDocumentHandler handler, string path)
    {
        if (handler is not OfficeCli.Handlers.PowerPointHandler) return null;
        try
        {
            var node = handler.Get(path);
            if (node == null) return null;
            // Only for spatial types (shape, textbox, picture, table, chart, connector, group, equation)
            if (node.Type is "slide" or "paragraph" or "run" or "cell" or "row") return null;
            if (!node.Format.ContainsKey("x") || !node.Format.ContainsKey("y")) return null;
            var x = node.Format.TryGetValue("x", out var xv) ? xv : "?";
            var y = node.Format.TryGetValue("y", out var yv) ? yv : "?";
            var w = node.Format.TryGetValue("width", out var wv) ? wv : "?";
            var h = node.Format.TryGetValue("height", out var hv) ? hv : "?";
            return $"x: {x}  y: {y}  width: {w}  height: {h}";
        }
        catch { return null; }
    }

    /// <summary>
    /// Check if the element at <paramref name="path"/> has the same (x,y) as any sibling.
    /// Returns list of overlapping sibling names, or empty.
    /// </summary>
    private static List<string> CheckPositionOverlap(IDocumentHandler handler, string path)
    {
        var overlaps = new List<string>();
        if (handler is not OfficeCli.Handlers.PowerPointHandler) return overlaps;
        try
        {
            var node = handler.Get(path);
            if (node == null || !node.Format.ContainsKey("x") || !node.Format.ContainsKey("y")) return overlaps;
            var myX = node.Format["x"]?.ToString();
            var myY = node.Format["y"]?.ToString();
            if (myX == null || myY == null) return overlaps;

            // Get parent (slide) to enumerate siblings
            var slidePathMatch = System.Text.RegularExpressions.Regex.Match(path, @"^(/slide\[\d+\])");
            if (!slidePathMatch.Success) return overlaps;
            var slidePath = slidePathMatch.Value;
            var slideNode = handler.Get(slidePath);
            if (slideNode == null) return overlaps;

            foreach (var child in slideNode.Children)
            {
                // Skip the element itself. `path` may be an index form
                // (/slide[1]/group[1]) while Get returns the canonical id form
                // (/slide[1]/group[@id=100000]); compare against BOTH so an
                // element never reports overlapping with itself.
                if (child.Path == path || child.Path == node.Path) continue;
                if (!child.Format.ContainsKey("x") || !child.Format.ContainsKey("y")) continue;
                var cx = child.Format["x"]?.ToString();
                var cy = child.Format["y"]?.ToString();
                if (cx == myX && cy == myY)
                {
                    // Skip false positive: both shapes at default (0,0) means neither was explicitly positioned
                    if (myX == "0cm" && myY == "0cm" && cx == "0cm" && cy == "0cm") continue;
                    var name = child.Format.TryGetValue("name", out var n) ? n?.ToString() : child.Path;
                    overlaps.Add(name ?? child.Path);
                }
            }
        }
        catch { /* ignore */ }
        return overlaps;
    }

    /// <summary>
    /// Check if a shape's text overflows its bounds using CJK-aware character measurement.
    /// Returns a warning message or null.
    /// </summary>
    internal static string? CheckTextOverflow(IDocumentHandler handler, string path)
    {
        try
        {
            return handler switch
            {
                OfficeCli.Handlers.PowerPointHandler ppt => ppt.CheckShapeTextOverflow(path),
                OfficeCli.Handlers.ExcelHandler xl => xl.CheckCellOverflow(path),
                _ => null
            };
        }
        catch { return null; }
    }

    /// <summary>
    /// Notify watch server with pre-rendered HTML from the handler.
    /// Call this while the handler is still open (before Dispose).
    /// </summary>
    private static void NotifyWatch(IDocumentHandler handler, string filePath, string? changedPath)
    {
        if (!WatchServer.IsWatching(filePath)) return;

        if (handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            string? scrollTo = null;
            var sheetName = WatchMessage.ExtractSheetName(changedPath);
            if (sheetName != null)
            {
                var idx = excel.GetSheetIndex(sheetName);
                if (idx >= 0) scrollTo = $".sheet-content[data-sheet=\"{idx}\"]";
            }
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = excel.ViewAsHtml(), ScrollTo = scrollTo });
            return;
        }
        if (handler is OfficeCli.Handlers.WordHandler word)
        {
            var scrollTo = WatchMessage.ExtractWordScrollTarget(changedPath);
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = word.ViewAsHtml(), ScrollTo = scrollTo });
            return;
        }
        if (handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var slideNum = WatchMessage.ExtractSlideNum(changedPath);
        if (slideNum > 0)
        {
            var html = ppt.RenderSlideHtml(slideNum);
            if (html != null)
            {
                // Slide-scoped replace: the watch server patches its cached _currentHtml in
                // place via PatchSlideInHtml; bundling a full ViewAsHtml() here is redundant
                // (and ResidentServer.NotifyWatchSlideChanged already omits it).
                WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "replace", Slide = slideNum, Html = html });
                return;
            }
        }
        WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = ppt.ViewAsHtml() });
    }

    private static void NotifyWatchRoot(IDocumentHandler handler, string filePath, int oldSlideCount)
    {
        if (!WatchServer.IsWatching(filePath)) return;

        if (handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = excel.ViewAsHtml() });
            return;
        }
        if (handler is OfficeCli.Handlers.WordHandler word)
        {
            // Scroll to last page (new content is typically appended)
            var html = word.ViewAsHtml();
            var pageCount = System.Text.RegularExpressions.Regex.Matches(html, @"data-page=""\d+""").Count;
            var scrollTo = pageCount > 0 ? $".page[data-page=\"{pageCount}\"]" : null;
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = html, ScrollTo = scrollTo });
            return;
        }
        if (handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var newCount = ppt.GetSlideCount();
        if (newCount > oldSlideCount)
        {
            var html = ppt.RenderSlideHtml(newCount);
            if (html != null)
            {
                WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "add", Slide = newCount, Html = html, FullHtml = ppt.ViewAsHtml() });
                return;
            }
        }
        else if (newCount < oldSlideCount)
        {
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "remove", Slide = oldSlideCount, FullHtml = ppt.ViewAsHtml() });
            return;
        }
        WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = ppt.ViewAsHtml() });
    }

    /// <summary>
    /// TextWriter that writes to two targets simultaneously (tee pattern).
    /// </summary>
    private class TeeWriter : TextWriter
    {
        private readonly TextWriter _a;
        private readonly TextWriter _b;
        public TeeWriter(TextWriter a, TextWriter b) { _a = a; _b = b; }
        public override Encoding Encoding => _a.Encoding;
        public override void Write(char value) { _a.Write(value); _b.Write(value); }
        public override void Write(string? value) { _a.Write(value); _b.Write(value); }
        public override void WriteLine(string? value) { _a.WriteLine(value); _b.WriteLine(value); }
        public override void Flush() { _a.Flush(); _b.Flush(); }
    }
}
