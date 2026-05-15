// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static readonly HashSet<string> HwpNativeReadOperations = new(StringComparer.OrdinalIgnoreCase)
    {
        "get-paragraph-count",
        "get-paragraph-length",
        "get-text-range",
        "get-style-list",
        "get-style",
        "get-numbering-list",
        "get-numbering",
        "get-bullet-list",
        "get-page-hide",
        "get-header",
        "get-footer",
        "get-picture",
        "get-shape",
        "get-equation",
        "get-footnote"
    };

    private static string? HwpViewOperationForMode(string modeKey)
        => modeKey is "text" or "t" ? HwpCapabilityConstants.OperationReadText
            : modeKey is "svg" or "g" ? HwpCapabilityConstants.OperationRenderSvg
            : modeKey is "png" ? HwpCapabilityConstants.OperationRenderPng
            : modeKey is "pdf" ? HwpCapabilityConstants.OperationExportPdf
            : modeKey is "markdown" or "md" ? HwpCapabilityConstants.OperationExportMarkdown
            : modeKey is "thumbnail" ? HwpCapabilityConstants.OperationThumbnail
            : modeKey is "info" ? HwpCapabilityConstants.OperationDocumentInfo
            : modeKey is "diagnostics" or "diag" ? HwpCapabilityConstants.OperationDiagnostics
            : modeKey is "dump" or "controls" ? HwpCapabilityConstants.OperationDumpControls
            : modeKey is "pages" or "dump-pages" ? HwpCapabilityConstants.OperationDumpPages
            : modeKey is "fields" ? HwpCapabilityConstants.OperationListFields
            : modeKey is "field" ? HwpCapabilityConstants.OperationReadField
            : modeKey is "table-cell" or "cell" ? HwpCapabilityConstants.OperationReadTableCell
            : modeKey is "tables" or "cells" ? HwpCapabilityConstants.OperationScanCells
            : modeKey is "native" or "native-op" ? HwpCapabilityConstants.OperationNativeRead
            : null;

    private static bool IsHwpNativeReadOperation(string op)
        => HwpNativeReadOperations.Contains(op);

    private static void ValidateHwpNativeViewRequest(
        string formatKey,
        string nativeOp,
        string[] nativeArgs)
    {
        if (!IsHwpNativeReadOperation(nativeOp))
            throw new HwpEngineException(
                $"HWP native view operation '{nativeOp}' is not read-only.",
                HwpCapabilityConstants.ReasonUnsupportedOperation,
                "Use `officecli set file.hwp /native-op --prop op=<op> --prop output=<path> ...` for native mutations.",
                [HwpCapabilityConstants.OperationNativeRead, HwpCapabilityConstants.OperationNativeMutation],
                formatKey,
                HwpCapabilityConstants.OperationNativeRead,
                HwpCapabilityConstants.EngineRhwpBridge,
                HwpCapabilityConstants.ModeExperimental);

        foreach (var (key, _) in ParsePropsArray(nativeArgs))
        {
            var normalized = key.TrimStart('-');
            if (normalized.Equals("output", StringComparison.OrdinalIgnoreCase)
                || normalized.Equals("out", StringComparison.OrdinalIgnoreCase))
                throw new HwpEngineException(
                    "HWP native view does not accept output paths.",
                    HwpCapabilityConstants.ReasonUnsupportedOperation,
                    "Use `officecli set file.hwp /native-op --prop op=<op> --prop output=<path> ...` for output-first native mutations.",
                    [HwpCapabilityConstants.OperationNativeRead, HwpCapabilityConstants.OperationNativeMutation],
                    formatKey,
                    HwpCapabilityConstants.OperationNativeRead,
                    HwpCapabilityConstants.EngineRhwpBridge,
                    HwpCapabilityConstants.ModeExperimental);
        }
    }
}
