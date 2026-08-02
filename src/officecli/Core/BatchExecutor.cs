// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Public entry point for running a batch against an already-open
/// <see cref="IDocumentHandler"/> without going through a CLI process or
/// resident pipe — for embedders that link officecli directly (in-process
/// hosts). Reuses the same batch-item dispatch and JSON
/// envelope shape as the CLI `batch` command and the Node/Python SDKs, so a
/// caller sees one consistent protocol regardless of transport.
/// </summary>
public static class BatchExecutor
{
    /// <summary>
    /// Runs a batch and returns EXACTLY what the CLI `batch` command writes to
    /// stdout, so an in-process host is byte-identical to the CLI (and to
    /// what the watch server relays): with <paramref name="json"/> the
    /// `{success, data:{results, summary}}` envelope, without it the
    /// `[i] ...\n\nBatch complete: ...` text. This is the same output path the
    /// CLI command uses (RunNonResidentBatch → PrintBatchResults → WrapEnvelope,
    /// see CommandBuilder.Batch.cs); the earlier shortcut here emitted a bare
    /// results array with `json` hardcoded, diverging from every other transport.
    /// </summary>
    /// <param name="itemsJson">A JSON array of batch items — the same shape documented
    /// for the CLI `batch` command and the `BatchItem` SDK type.</param>
    /// <param name="json">Mirrors the CLI `--json` toggle: structured envelope vs plain text.</param>
    public static string ExecuteBatch(IDocumentHandler handler, string itemsJson, bool json, bool stopOnError = false)
    {
        try
        {
            var items = JsonSerializer.Deserialize(itemsJson, BatchJsonContext.Default.ListBatchItem) ?? new List<BatchItem>();
            // NEWLINE-SEMANTICS-V2: strip meta items; rewrite legacy docx dumps.
            BatchCompat.PrepareForReplay(items,
                handler is OfficeCli.Handlers.WordHandler ? ".docx" : "");
            var unrecognizedLatex = new List<string>();
            var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError, json, unrecognizedLatex);

            // Reuse the CLI's own body formatter so both modes are byte-identical.
            using var sw = new System.IO.StringWriter();
            CommandBuilder.PrintBatchResults(results, json, items.Count, sw);
            var inner = sw.ToString().TrimEnd('\n', '\r');
            if (!json) return inner;

            // Same outer envelope the CLI batch command applies: unrecognized-LaTeX
            // warnings + outer success true only when every step succeeded
            // (CommandBuilder.Batch.cs). Judgment command — a single failed step
            // flips outer success to false.
            var warnings = new List<CliWarning>();
            foreach (var tok in unrecognizedLatex)
                warnings.Add(new CliWarning
                {
                    Message = $"unrecognized_latex_command: {tok}",
                    Code = "unrecognized_latex_command",
                    Suggestion = "Check the command spelling; see https://katex.org/docs/supported.html for supported syntax.",
                });
            var success = results.Count == 0 || !results.Any(r => !r.Success);
            return OutputFormatter.WrapEnvelope(inner, warnings, success);
        }
        catch (Exception ex)
        {
            return RenderTopLevelError(ex, json);
        }
    }

    /// <summary>
    /// Mirrors the SDKs' <c>send(item)</c> (as distinct from <c>batch(items)</c>):
    /// runs ONE batch-shaped item and returns EXACTLY what the CLI single command
    /// writes to stdout — with <paramref name="json"/> the standalone command's
    /// envelope (WrapEnvelopeText's `{success,data,message}` for text commands
    /// like set/add/remove, WrapEnvelope's `{success,data}` for data commands like
    /// get/query), without it the plain text. A failure throws instead of being
    /// captured per-item, matching the standalone command's exit contract.
    /// </summary>
    /// <param name="itemJson">A single batch item object — the same shape as one
    /// entry in <see cref="ExecuteBatch"/>'s array, or the SDKs' `send(item)` argument.</param>
    /// <param name="json">Mirrors the CLI `--json` toggle: structured envelope vs plain text.</param>
    public static string ExecuteSend(IDocumentHandler handler, string itemJson, bool json)
    {
        try
        {
            var item = JsonSerializer.Deserialize(itemJson, BatchJsonContext.Default.BatchItem)
                ?? throw new ArgumentException("send: empty item");
            var inner = CommandBuilder.ExecuteBatchItem(handler, item, json);
            if (!json) return inner;

            // Match the standalone command's envelope by inner shape — a JSON payload
            // (get/query/view/validate/…) wraps with WrapEnvelope; a plain-text
            // message (set/add/remove/…) wraps with WrapEnvelopeText (which adds the
            // `message` field the CLI emits). Content-based so a new command inherits
            // the right envelope without a hand-kept command→envelope map.
            var trimmed = inner.TrimStart();
            return trimmed.StartsWith('{') || trimmed.StartsWith('[')
                ? OutputFormatter.WrapEnvelope(inner)
                : OutputFormatter.WrapEnvelopeText(inner);
        }
        catch (Exception ex)
        {
            return RenderTopLevelError(ex, json);
        }
    }

    /// <summary>
    /// Mirror of the CLI's shared <c>WriteError(ex, json)</c> (CommandBuilder):
    /// in JSON mode a failed command puts the <c>{success:false, error:{…}}</c>
    /// envelope on stdout — return it byte-identically. In text mode the CLI
    /// writes to stderr and leaves stdout empty; with no stderr channel in-process
    /// we surface the error by propagating it (the embedder host's failure
    /// signal), matching the prior throw-on-failure contract.
    /// </summary>
    private static string RenderTopLevelError(Exception ex, bool json)
    {
        // CONSISTENCY(error-wrap): same friendlier rendering the CLI applies to a
        // bare XmlException from an externally-corrupted OOXML part.
        var rendered = ex is System.Xml.XmlException xe
            ? new System.IO.InvalidDataException(
                $"Malformed XML in document part: {xe.Message} " +
                $"(the file appears to have a corrupted OOXML part).", xe)
            : ex;
        if (json) return OutputFormatter.WrapErrorEnvelope(rendered);
        throw rendered;
    }
}
