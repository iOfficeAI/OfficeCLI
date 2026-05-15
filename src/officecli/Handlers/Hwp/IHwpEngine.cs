// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp;

public interface IHwpEngine
{
    string Name { get; }
    string? Version { get; }
    HwpEngineMode Mode { get; }

    Task<HwpCapabilityReport> GetCapabilitiesAsync(CancellationToken ct);
    Task<HwpTextResult> ReadTextAsync(HwpReadRequest request, CancellationToken ct);
    Task<HwpRenderResult> RenderSvgAsync(HwpRenderRequest request, CancellationToken ct);
    Task<HwpJsonViewResult> ViewJsonAsync(HwpJsonViewRequest request, CancellationToken ct);
    Task<HwpFieldListResult> ListFieldsAsync(HwpFieldListRequest request, CancellationToken ct);
    Task<HwpFieldReadResult> ReadFieldAsync(HwpFieldReadRequest request, CancellationToken ct);
    Task<HwpMutationResult> FillFieldAsync(HwpFillFieldRequest request, CancellationToken ct);
    Task<HwpMutationResult> ReplaceTextAsync(HwpReplaceTextRequest request, CancellationToken ct);
    Task<HwpMutationResult> InsertTextAsync(HwpInsertTextRequest request, CancellationToken ct);
    Task<HwpMutationResult> SetTableCellAsync(HwpTableCellSetRequest request, CancellationToken ct);
    Task<HwpMutationResult> SaveOriginalAsync(HwpSaveOriginalRequest request, CancellationToken ct);
    Task<HwpMutationResult> ConvertToEditableAsync(HwpConvertToEditableRequest request, CancellationToken ct);
    Task<HwpMutationResult> NativeMutationAsync(HwpNativeMutationRequest request, CancellationToken ct);
    Task<HwpMutationResult> SaveAsHwpAsync(HwpSaveAsHwpRequest request, CancellationToken ct);
}
