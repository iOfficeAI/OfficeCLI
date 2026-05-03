// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Handlers.Hwp;

public sealed class CustomHwpxEngine : IHwpEngine
{
    public string Name => HwpCapabilityConstants.EngineCustom;
    public string? Version => $"officecli:{Assembly.GetExecutingAssembly().GetName().Version}";
    public HwpEngineMode Mode => HwpEngineMode.Default;

    public Task<HwpCapabilityReport> GetCapabilitiesAsync(CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        return Task.FromResult(HwpCapabilityFactory.BuildReport(Version));
    }

    public Task<HwpTextResult> ReadTextAsync(HwpReadRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationReadText);
    }

    public Task<HwpRenderResult> RenderSvgAsync(HwpRenderRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationRenderSvg);
    }

    public Task<HwpFieldListResult> ListFieldsAsync(HwpFieldListRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationListFields);
    }

    public Task<HwpFieldReadResult> ReadFieldAsync(HwpFieldReadRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationReadField);
    }

    public Task<HwpMutationResult> FillFieldAsync(HwpFillFieldRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationFillField);
    }

    public Task<HwpMutationResult> ReplaceTextAsync(HwpReplaceTextRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationReplaceText);
    }

    public Task<HwpMutationResult> SaveOriginalAsync(HwpSaveOriginalRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationSaveOriginal);
    }

    public Task<HwpMutationResult> SaveAsHwpAsync(HwpSaveAsHwpRequest request, CancellationToken ct)
    {
        throw Unsupported(request.Format, HwpCapabilityConstants.OperationSaveAsHwp);
    }

    private static HwpEngineException Unsupported(HwpFormat format, string operation)
    {
        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
        var reason = operation switch
        {
            HwpCapabilityConstants.OperationFillField when format == HwpFormat.Hwp
                => HwpCapabilityConstants.ReasonBinaryHwpMutationForbidden,
            HwpCapabilityConstants.OperationSaveOriginal or HwpCapabilityConstants.OperationSaveAsHwp when format == HwpFormat.Hwp
                => HwpCapabilityConstants.ReasonBinaryHwpWriteForbidden,
            _ => HwpCapabilityConstants.ReasonRoundTripUnverified
        };

        return new HwpEngineException(
            $"{formatKey} operation '{operation}' is not roundtrip-verified in this OfficeCLI build.",
            reason,
            "Check `officecli capabilities --json` and use only roundtrip-verified operations.",
            [
                HwpCapabilityConstants.OperationReadText,
                HwpCapabilityConstants.OperationRenderSvg,
                HwpCapabilityConstants.OperationListFields,
                HwpCapabilityConstants.OperationReadField,
                HwpCapabilityConstants.OperationFillField,
                HwpCapabilityConstants.OperationReplaceText,
                HwpCapabilityConstants.OperationSaveOriginal,
                HwpCapabilityConstants.OperationSaveAsHwp
            ],
            formatKey,
            operation,
            HwpCapabilityConstants.EngineCustom,
            HwpCapabilityConstants.ModeDefault);
    }
}
