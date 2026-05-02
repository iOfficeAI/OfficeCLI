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
    Task<HwpMutationResult> FillFieldAsync(HwpFillFieldRequest request, CancellationToken ct);
    Task<HwpMutationResult> SaveOriginalAsync(HwpSaveOriginalRequest request, CancellationToken ct);
    Task<HwpMutationResult> SaveAsHwpAsync(HwpSaveAsHwpRequest request, CancellationToken ct);
}
