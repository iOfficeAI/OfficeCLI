// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp;

using System.Text.Json.Nodes;

public sealed record OfficeCliJsonEnvelope<T>(
    bool Success,
    T? Data,
    string[] Warnings,
    OfficeCliError? Error
);

public sealed record OfficeCliError(
    string Error,
    string Code,
    string? Suggestion,
    string[] ValidValues,
    string? Format,
    string? Operation,
    string? Engine,
    string? EngineMode
);

public sealed class HwpEngineException : Exception
{
    public HwpEngineException(
        string message,
        string code,
        string? suggestion = null,
        string[]? validValues = null,
        string? format = null,
        string? operation = null,
        string? engine = null,
        string? engineMode = null,
        JsonObject? transaction = null) : base(message)
    {
        Error = new OfficeCliError(
            message,
            code,
            suggestion,
            validValues ?? [],
            format,
            operation,
            engine,
            engineMode);
        Transaction = transaction;
    }

    public OfficeCliError Error { get; }
    public JsonObject? Transaction { get; }
}
