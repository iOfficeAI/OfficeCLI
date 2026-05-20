// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddSlide(string parentPath, int? index, Dictionary<string, string> properties)
    {
                properties ??= new Dictionary<string, string>();
                // A slide can only attach to the presentation root. Earlier
                // releases silently fell back to "/" when the caller passed
                // a non-root parent (e.g. `/slide[1]`, `/section[2]`,
                // `/bogus`), which masked path typos and produced a slide
                // somewhere other than where the caller expected.
                if (!string.IsNullOrEmpty(parentPath)
                    && parentPath != "/"
                    && parentPath != "")
                {
                    throw new ArgumentException(
                        $"Invalid parent '{parentPath}' for --type slide: a slide can only be added at '/' " +
                        "(slides hang off the presentation root, not under another element).");
                }
                var presentationPart = _doc.PresentationPart
                    ?? throw new InvalidOperationException("Presentation not found");
                var presentation = presentationPart.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var slideIdList = presentation.GetFirstChild<SlideIdList>()
                    ?? presentation.AppendChild(new SlideIdList());

                var newSlidePart = presentationPart.AddNewPart<SlidePart>();

                // Link slide to slideLayout (required by PowerPoint)
                var slideLayoutPart = ResolveSlideLayout(
                    presentationPart, properties.GetValueOrDefault("layout"));
                if (slideLayoutPart != null)
                    newSlidePart.AddPart(slideLayoutPart);

                newSlidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties()
                        )
                    )
                );

                // Add title shape if text provided (ID starts at 2 since ShapeTree group uses ID=1)
                uint nextShapeId = 2;
                if (properties.TryGetValue("title", out var titleText))
                {
                    XmlTextValidator.ValidateOrThrow(titleText, "title");
                    var titleShape = CreateTextShape(nextShapeId++, "Title", titleText, true);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(titleShape);
                }

                // Add content text if provided
                if (properties.TryGetValue("text", out var contentText))
                {
                    XmlTextValidator.ValidateOrThrow(contentText, "text");
                    // Symmetry with the title path above: title carries
                    // <p:ph type="title"/>, so content carries
                    // <p:ph type="body" idx="1"/> — both bind to layout
                    // slots and Get reports them as placeholder-flavored
                    // (title → type=title; content → type=placeholder +
                    // phType=body) instead of mismatched title vs bare textbox.
                    var textShape = CreateTextShape(nextShapeId++, "Content", contentText, false, isTextBox: true,
                        placeholderType: PlaceholderValues.Body, placeholderIndex: 1);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(textShape);
                }

                // Apply background if provided
                if (properties.TryGetValue("background", out var bgValue))
                    ApplySlideBackground(newSlidePart, bgValue);

                // Apply transition if provided
                if (properties.TryGetValue("transition", out var transValue))
                {
                    ApplyTransition(newSlidePart, transValue);
                    if (transValue.StartsWith("morph", StringComparison.OrdinalIgnoreCase))
                        AutoPrefixMorphNames(newSlidePart);
                }
                if (properties.TryGetValue("advancetime", out var advTime) || properties.TryGetValue("advanceTime", out advTime))
                    SetAdvanceTime(newSlidePart.Slide, advTime);
                if (properties.TryGetValue("advanceclick", out var advClick) || properties.TryGetValue("advanceClick", out advClick))
                    SetAdvanceClick(newSlidePart.Slide, IsTruthy(advClick));
                if (properties.TryGetValue("hidden", out var hiddenVal) && IsTruthy(hiddenVal))
                    newSlidePart.Slide.Show = false;

                newSlidePart.Slide.Save();

                var maxId = slideIdList.Elements<SlideId>().Any()
                    ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
                    : 256;
                var relId = presentationPart.GetIdOfPart(newSlidePart);

                if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
                {
                    var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
                    if (refSlide != null)
                        slideIdList.InsertBefore(new SlideId { Id = maxId, RelationshipId = relId }, refSlide);
                    else
                        slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }
                else
                {
                    slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }

                presentation.Save();
                // Find the actual position of the inserted slide
                var slideIds = slideIdList.Elements<SlideId>().ToList();
                var insertedIdx = slideIds.FindIndex(s => s.RelationshipId?.Value == relId) + 1;
                return $"/slide[{insertedIdx}]";
    }


}
