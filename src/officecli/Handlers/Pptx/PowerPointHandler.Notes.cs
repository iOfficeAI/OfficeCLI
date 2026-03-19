// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Speaker Notes helpers ====================

    private static string GetNotesText(NotesSlidePart notesPart)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return "";

        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1) // body/notes placeholder
            {
                return string.Join("\n", shape.TextBody?.Elements<Drawing.Paragraph>()
                    .Select(p => string.Concat(p.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? "")))
                    ?? Enumerable.Empty<string>());
            }
        }
        return "";
    }

    private static void SetNotesText(NotesSlidePart notesPart, string text)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Notes slide has no shape tree");

        // Find body placeholder (idx=1)
        Shape? notesShape = null;
        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1)
            {
                notesShape = shape;
                break;
            }
        }

        if (notesShape == null)
        {
            notesShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 3, Name = "Notes Placeholder 2" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 }
                    )
                ),
                new ShapeProperties(),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );
            spTree.AppendChild(notesShape);
        }

        var textBody = notesShape.TextBody
            ?? (notesShape.TextBody = new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle()));

        textBody.RemoveAllChildren<Drawing.Paragraph>();
        foreach (var line in text.Split('\n'))
        {
            textBody.AppendChild(new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text(line)
                )
            ));
        }

        notesPart.NotesSlide!.Save();
    }

    private static NotesSlidePart EnsureNotesSlidePart(SlidePart slidePart)
    {
        if (slidePart.NotesSlidePart != null) return slidePart.NotesSlidePart;

        var notesPart = slidePart.AddNewPart<NotesSlidePart>();
        notesPart.NotesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(new Drawing.TransformGroup()),
                    // Slide image placeholder (idx=0)
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 2, Name = "Slide Image Placeholder 1" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.SlideImage, Index = 0 }
                            )
                        ),
                        new ShapeProperties(),
                        new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
                    ),
                    // Notes body placeholder (idx=1)
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 3, Name = "Notes Placeholder 2" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 }
                            )
                        ),
                        new ShapeProperties(),
                        new TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph()
                        )
                    )
                )
            )
        );
        notesPart.NotesSlide.Save();
        return notesPart;
    }
}
