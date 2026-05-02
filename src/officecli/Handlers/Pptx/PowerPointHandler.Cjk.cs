// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private static List<Drawing.Run> BuildSegmentedRuns(string text, Drawing.RunProperties? template = null, string fallbackLang = "en-US")
    {
        var segments = CjkHelper.SegmentText(text);
        if (segments.Count == 0)
            segments = new List<(string text, CjkScript script)> { ("", CjkScript.None) };

        var runs = new List<Drawing.Run>();
        foreach (var (segmentText, script) in segments)
        {
            var rPr = template?.CloneNode(true) as Drawing.RunProperties ?? new Drawing.RunProperties();
            if (script != CjkScript.None)
                CjkHelper.ApplyToDrawingRun(rPr, script);
            else
                CjkHelper.ClearDrawingRunCjk(rPr, fallbackLang);

            runs.Add(new Drawing.Run(
                rPr,
                new Drawing.Text { Text = segmentText }));
        }

        return runs;
    }

    private static Drawing.Paragraph BuildParagraphWithSegmentedRuns(
        string text,
        Drawing.RunProperties? template = null,
        Drawing.ParagraphProperties? paragraphProperties = null)
    {
        var paragraph = new Drawing.Paragraph();
        if (paragraphProperties != null)
            paragraph.ParagraphProperties = paragraphProperties.CloneNode(true) as Drawing.ParagraphProperties;

        foreach (var run in BuildSegmentedRuns(text, template))
            paragraph.Append(run);

        return paragraph;
    }

    private static void ReplaceRunWithSegmentedRuns(Drawing.Run run, string text)
    {
        if (run.Parent is not Drawing.Paragraph paragraph)
        {
            run.Text = new Drawing.Text { Text = text };
            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
            CjkHelper.ApplyToDrawingRunIfCjk(rPr, text);
            return;
        }

        var template = run.RunProperties?.CloneNode(true) as Drawing.RunProperties;
        foreach (var newRun in BuildSegmentedRuns(text, template))
            paragraph.InsertBefore(newRun, run);

        run.Remove();
    }
}
