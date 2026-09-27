// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Runtime.InteropServices;

namespace OfficeCli.Core;

/// <summary>
/// Converts native Office rendering failures into actionable CLI diagnostics.
/// Auto rendering still falls back silently; this is used only when the caller
/// explicitly requested the native backend.
/// </summary>
internal static class NativeRenderDiagnostics
{
    const int ClassNotRegistered = unchecked((int)0x80040154);

    internal static CliException Create(
        string application,
        bool attempted,
        Exception? failure)
    {
        if (!attempted)
        {
            return new CliException($"--render native requires Windows with {application} installed.")
            {
                Code = "native_unavailable",
                Suggestion = "Use --render html or --render auto."
            };
        }

        if (failure == null)
        {
            return new CliException($"{application} native render did not produce an image.")
            {
                Code = "native_render_failed",
                Suggestion = "Verify the document and page range, or use --render html or --render auto."
            };
        }

        var detail = Describe(failure);
        if (IsUnavailable(failure))
        {
            return new CliException($"{application} is unavailable: {detail}", failure)
            {
                Code = "native_unavailable",
                Suggestion = $"Install or repair {application}, or use --render html or --render auto."
            };
        }

        return new CliException($"{application} native render failed: {detail}", failure)
        {
            Code = "native_render_failed",
            Suggestion = "Verify that the document opens in Office, or use --render html or --render auto."
        };
    }

    static bool IsUnavailable(Exception failure)
        => failure is COMException { HResult: ClassNotRegistered }
           || failure.Message.StartsWith("app_not_authentic:", StringComparison.Ordinal)
           || failure.Message.StartsWith("word_not_authentic:", StringComparison.Ordinal);

    static string Describe(Exception failure)
    {
        var message = string.IsNullOrWhiteSpace(failure.Message)
            ? failure.GetType().Name
            : failure.Message.Trim();
        if (failure is ExternalException && !message.Contains("0x", StringComparison.OrdinalIgnoreCase))
            message += $" (HRESULT 0x{failure.HResult:X8})";
        return message;
    }
}
