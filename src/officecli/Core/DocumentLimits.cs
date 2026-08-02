// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Runtime.CompilerServices;

namespace OfficeCli.Core;

/// <summary>
/// CONSISTENCY(dos-hardening): one source of truth for the resource limits that
/// protect officecli from hostile/malformed documents. These guard the three
/// denial-of-service classes a tiny crafted file can trigger:
///
///  - Decompression bombs: an OOXML zip whose entries inflate to gigabytes.
///    Enforced in <see cref="OfficeCli.Handlers.DocumentHandlerFactory"/> before
///    the Open XML SDK touches the package.
///  - Unbounded structural recursion: deeply nested tables / group shapes drive
///    the tree walkers and HTML/SVG renderers into an uncatchable
///    StackOverflowException (which escapes the top-level SafeRun handler and
///    hard-kills the process — fatal in the long-lived resident/watch servers).
///    Each recursive walker checks <see cref="MaxRecursionDepth"/> and throws a
///    friendly <see cref="CliException"/> instead.
///  - Catastrophic-regex backtracking: a user-supplied r"..." pattern against
///    document text. Bounded by <see cref="RegexMatchTimeout"/>, mirroring the
///    existing guard in <see cref="FindHelpers"/>.
///
/// Limits are deliberately generous — far beyond any legitimate document — so
/// real files are never affected; only adversarial inputs hit them.
/// </summary>
public static class DocumentLimits
{
    /// <summary>
    /// Maximum element nesting depth any recursive document walker / renderer
    /// will descend before refusing. Real Office documents nest only a handful
    /// of levels deep (Word caps nested tables at ~19; PPTX group nesting and
    /// math run far lower), so 256 is comfortably above any legitimate file yet
    /// low enough that even the heavy HTML/SVG renderer frames cannot overflow a
    /// default ~1 MB thread-pool stack (the resident/watch server runs commands
    /// on such threads). The <see cref="RuntimeHelpers"/> stack probe in
    /// <see cref="EnsureDepth"/> backs this up for any unusually large frame.
    /// </summary>
    public const int MaxRecursionDepth = 256;

    /// <summary>
    /// Maximum total uncompressed size (bytes) of all entries in an OOXML
    /// package. 2 GiB matches realistic large-but-legitimate documents (big
    /// embedded media) while rejecting decompression bombs that inflate a few
    /// KB of zip into many gigabytes.
    /// </summary>
    public const long MaxUncompressedBytes = 2L * 1024 * 1024 * 1024;

    /// <summary>
    /// Maximum number of entries in an OOXML package. Guards against zip files
    /// crafted with millions of tiny entries (entry-count exhaustion).
    /// </summary>
    public const int MaxZipEntries = 100_000;

    /// <summary>
    /// Maximum overall compression ratio (uncompressed / compressed) tolerated
    /// for a package. Genuine OOXML packages rarely exceed ~100×; a ratio far
    /// above this is the signature of a decompression bomb.
    /// </summary>
    public const long MaxCompressionRatio = 1000;

    /// <summary>
    /// Maximum number of XML elements officecli will let a single package
    /// materialize into an in-memory DOM. The zip guards above bound *bytes*, but
    /// the Open XML SDK builds a full node tree on any read that walks it, and a
    /// worksheet of millions of tiny <c><v>…</v></c> cells costs ~hundreds of
    /// bytes of managed heap per cell — a few MB of zip (well under every byte /
    /// ratio / entry guard) expands into multiple GiB of RAM and OOM-kills the
    /// long-lived resident/watch server. This cap is checked with a streaming
    /// pre-scan (XmlReader, no DOM) so it trips before the expensive
    /// materialization. Generous — a legitimate workbook with a few hundred
    /// thousand populated cells stays far below it — and env-overridable for the
    /// rare genuinely enormous data workbook, mirroring the sweep-budget knob.
    /// </summary>
    public static readonly long MaxDomElements =
        long.TryParse(Environment.GetEnvironmentVariable("OFFICECLI_MAX_DOM_ELEMENTS"),
            out var e) && e > 0 ? e : 3_000_000;

    /// <summary>
    /// Only parts whose uncompressed size exceeds this are stream-counted against
    /// <see cref="MaxDomElements"/>. A normal document's parts are all far smaller,
    /// so the common case pays zero extra scan cost; only a suspiciously large
    /// part is walked, and that walk is cheap relative to the DOM build it guards.
    /// </summary>
    public const long ElementScanPartThreshold = 8L * 1024 * 1024;

    /// <summary>
    /// Hard timeout for matching a user-supplied regular expression against
    /// document text. Mirrors <see cref="FindHelpers.RegexMatchTimeout"/> so
    /// every find-style entry point fails fast on catastrophic backtracking
    /// instead of hanging the process.
    /// </summary>
    public static readonly TimeSpan RegexMatchTimeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Throw a friendly <see cref="CliException"/> when a recursive walker /
    /// renderer has descended too far. Call at the top of each recursive method
    /// so a maliciously deep document fails with a clean error instead of an
    /// uncatchable StackOverflowException.
    ///
    /// Two complementary guards, because the safe depth depends on thread stack
    /// size (the 8 MB main thread tolerates far deeper recursion than the ~1 MB
    /// thread-pool threads the resident/watch server uses, and renderer frames
    /// are large):
    ///  - <see cref="MaxRecursionDepth"/> bounds the worst-case time/O(n^2) cost
    ///    on any stack;
    ///  - <see cref="RuntimeHelpers.TryEnsureSufficientExecutionStack"/> probes
    ///    the *actual* remaining stack and trips before a real overflow, so the
    ///    guard adapts to whatever thread the call runs on (mirrors the probe in
    ///    <see cref="OfficeCli.Core.Formula.FormulaEvaluator"/>).
    /// </summary>
    public static void EnsureDepth(int depth)
    {
        if (depth > MaxRecursionDepth || !RuntimeHelpers.TryEnsureSufficientExecutionStack())
            throw new CliException(
                $"Document nesting exceeds the maximum supported depth (~{MaxRecursionDepth}); " +
                "the file may be malformed or crafted to exhaust resources.")
            {
                Code = "max_depth_exceeded",
                Suggestion = "Verify the document is a genuine Office file."
            };
    }
}
