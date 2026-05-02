// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private static List<Run> BuildSegmentedRuns(string text, RunProperties? template = null)
    {
        var segments = CjkHelper.SegmentText(text);
        if (segments.Count == 0)
            segments = new List<(string text, CjkScript script)> { ("", CjkScript.None) };

        var runs = new List<Run>();
        foreach (var (segmentText, script) in segments)
        {
            var rPr = template?.CloneNode(true) as RunProperties ?? new RunProperties();
            if (script != CjkScript.None)
                CjkHelper.ApplyToWordRun(rPr, script);
            else
                CjkHelper.ClearWordRunCjk(rPr);

            runs.Add(new Run(
                rPr,
                new Text(segmentText) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return runs;
    }
}
