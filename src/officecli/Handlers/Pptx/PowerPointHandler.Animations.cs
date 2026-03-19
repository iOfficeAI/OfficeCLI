// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Slide Transitions ====================

    /// <summary>
    /// Apply (or remove) a slide transition.
    /// Format: "TYPE[-DIR][-SPEED|DUR]" or "none"
    ///   TYPE: fade, cut, dissolve, wipe, push, cover, pull, split, zoom, wheel,
    ///         blinds, checker, comb, bars, strips, circle, diamond, newsflash,
    ///         plus, random, wedge, flash, honeycomb, vortex, switch, flip, ripple,
    ///         glitter, prism, doors, window, shred, ferris, flythrough, warp,
    ///         gallery, conveyor, pan, reveal
    ///   DIR: left/right/up/down (for wipe/push/cover/pull), in/out (for zoom/split)
    ///        horizontal/vertical/vert/horz (for blinds/checker/comb/bars/split)
    ///   SPEED: slow / medium|med / fast
    ///   DUR:   integer in ms (e.g. 1000) — requires Office 2010+
    /// Additional properties (set separately):
    ///   advancetime=3000    auto-advance after N ms
    ///   advanceclick=false  disable click-to-advance
    /// Examples: "fade", "wipe-left", "push-right", "split-horizontal-in", "zoom-out-slow", "none"
    /// </summary>
    private static void ApplyTransition(SlidePart slidePart, string value)
    {
        var slide = slidePart.Slide ?? throw new InvalidOperationException("Corrupt file");

        // Step 1: Build the Transition element using SDK (for correct child XML generation)
        var parts = value.Split('-');
        var typeName = parts[0].ToLowerInvariant();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            slide.Transition = null;
            return;
        }

        TransitionSpeedValues? speed = null;
        string? durationMs = null;
        string? direction = null;

        foreach (var part in parts.Skip(1))
        {
            var p = part.ToLowerInvariant();
            if (int.TryParse(p, out _))
                durationMs = p;
            else if (p == "slow")
                speed = TransitionSpeedValues.Slow;
            else if (p is "fast")
                speed = TransitionSpeedValues.Fast;
            else if (p is "medium" or "med")
                speed = TransitionSpeedValues.Medium;
            else
                direction = p;
        }

        var trans = new Transition();
        if (speed.HasValue) trans.Speed = speed.Value;
        if (durationMs != null) trans.Duration = durationMs;

        OpenXmlElement? transElem = typeName switch
        {
            "fade" => new FadeTransition(),
            "cut" => new CutTransition(),
            "dissolve" => new DissolveTransition(),
            "circle" => new CircleTransition(),
            "diamond" => new DiamondTransition(),
            "newsflash" => new NewsflashTransition(),
            "plus" => new PlusTransition(),
            "random" => new RandomTransition(),
            "wedge" => new WedgeTransition(),
            "wipe" => new WipeTransition { Direction = ParseSlideDir(direction ?? "left") },
            "push" => new PushTransition { Direction = ParseSlideDir(direction ?? "left") },
            "cover" => new CoverTransition { Direction = ParseSlideDirStr(direction ?? "left") },
            "pull" or "uncover" => new PullTransition { Direction = ParseSlideDirStr(direction ?? "right") },
            "wheel" => new WheelTransition { Spokes = new UInt32Value(4u) },
            "zoom" or "box" => new ZoomTransition { Direction = ParseInOutDir(direction ?? "in") },
            "split" => BuildSplitTransition(direction),
            "blinds" or "venetian" => new BlindsTransition { Direction = ParseOrientation(direction ?? "horizontal") },
            "checker" or "checkerboard" => new CheckerTransition { Direction = ParseOrientation(direction ?? "horizontal") },
            "comb" => new CombTransition { Direction = ParseOrientation(direction ?? "horizontal") },
            "bars" or "randombar" => new RandomBarTransition { Direction = ParseOrientation(direction ?? "horizontal") },
            "strips" or "diagonal" => new StripsTransition { Direction = ParseCornerDir(direction ?? "rd") },
            "flash" => new DocumentFormat.OpenXml.Office2010.PowerPoint.FlashTransition(),
            "honeycomb" => new DocumentFormat.OpenXml.Office2010.PowerPoint.HoneycombTransition(),
            "vortex" => new DocumentFormat.OpenXml.Office2010.PowerPoint.VortexTransition { Direction = ParseSlideDir(direction ?? "left") },
            "switch" => new DocumentFormat.OpenXml.Office2010.PowerPoint.SwitchTransition(),
            "flip" => new DocumentFormat.OpenXml.Office2010.PowerPoint.FlipTransition(),
            "ripple" => new DocumentFormat.OpenXml.Office2010.PowerPoint.RippleTransition(),
            "glitter" => new DocumentFormat.OpenXml.Office2010.PowerPoint.GlitterTransition { Direction = ParseSlideDir(direction ?? "left") },
            "prism" => new DocumentFormat.OpenXml.Office2010.PowerPoint.PrismTransition(),
            "doors" => new DocumentFormat.OpenXml.Office2010.PowerPoint.DoorsTransition { Direction = ParseOrientation(direction ?? "horizontal") },
            "window" => new DocumentFormat.OpenXml.Office2010.PowerPoint.WindowTransition { Direction = ParseOrientation(direction ?? "horizontal") },
            "shred" => new DocumentFormat.OpenXml.Office2010.PowerPoint.ShredTransition(),
            "ferris" => new DocumentFormat.OpenXml.Office2010.PowerPoint.FerrisTransition(),
            "flythrough" => new DocumentFormat.OpenXml.Office2010.PowerPoint.FlythroughTransition(),
            "warp" => new DocumentFormat.OpenXml.Office2010.PowerPoint.WarpTransition(),
            "gallery" => new DocumentFormat.OpenXml.Office2010.PowerPoint.GalleryTransition(),
            "conveyor" => new DocumentFormat.OpenXml.Office2010.PowerPoint.ConveyorTransition(),
            "pan" => new DocumentFormat.OpenXml.Office2010.PowerPoint.PanTransition { Direction = ParseSlideDir(direction ?? "left") },
            "reveal" => new DocumentFormat.OpenXml.Office2010.PowerPoint.RevealTransition(),
            "morph" => null, // handled specially below
            _ => throw new ArgumentException($"Invalid transition type: '{typeName}'. Valid values: fade, cut, dissolve, circle, diamond, newsflash, plus, random, wedge, wipe, push, cover, pull, wheel, zoom, split, blinds, checker, comb, bars, strips, flash, honeycomb, vortex, switch, flip, ripple, glitter, prism, doors, window, shred, ferris, flythrough, warp, gallery, conveyor, pan, reveal, morph, none.")
        };

        // Morph transition: requires mc:AlternateContent wrapper with p159 namespace
        // PowerPoint ignores bare p159:morph without the mc:Choice/Fallback structure
        if (typeName == "morph")
        {
            var morphOption = (direction ?? "byobject").ToLowerInvariant() switch
            {
                "byword" or "word" => "byWord",
                "bychar" or "char" or "character" => "byChar",
                "byobject" or "object" => "byObject",
                _ => throw new ArgumentException($"Invalid morph option: '{direction}'. Valid values: byObject, byWord, byChar.")
            };

            var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            var pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
            var p159Ns = "http://schemas.microsoft.com/office/powerpoint/2015/09/main";

            // Build speed/duration attributes
            var spdAttr = speed.HasValue ? $" spd=\"{((IEnumValue)speed.Value).Value}\"" : "";
            var durAttr = durationMs != null ? $" dur=\"{durationMs}\"" : "";

            // mc:AlternateContent > mc:Choice[Requires=p159] > p:transition > p159:morph
            var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);
            var choiceElement = new OpenXmlUnknownElement("mc", "Choice", mcNs);
            choiceElement.SetAttribute(new OpenXmlAttribute("", "Requires", null!, "p159"));

            var morphTrans = new OpenXmlUnknownElement("p", "transition", pNs);
            morphTrans.AddNamespaceDeclaration("p159", p159Ns);
            if (speed.HasValue)
                morphTrans.SetAttribute(new OpenXmlAttribute("", "spd", null!, ((IEnumValue)speed.Value).Value));
            if (durationMs != null)
                morphTrans.SetAttribute(new OpenXmlAttribute("", "dur", null!, durationMs));
            var morphElem = new OpenXmlUnknownElement("p159", "morph", p159Ns);
            morphElem.SetAttribute(new OpenXmlAttribute("", "option", null!, morphOption));
            morphTrans.AppendChild(morphElem);
            choiceElement.AppendChild(morphTrans);

            // mc:Fallback > p:transition > p:fade (graceful degradation for older PPT)
            var fallbackElement = new OpenXmlUnknownElement("mc", "Fallback", mcNs);
            var fallbackTrans = new OpenXmlUnknownElement("p", "transition", pNs);
            if (speed.HasValue)
                fallbackTrans.SetAttribute(new OpenXmlAttribute("", "spd", null!, ((IEnumValue)speed.Value).Value));
            fallbackTrans.AppendChild(new OpenXmlUnknownElement("p", "fade", pNs));
            fallbackElement.AppendChild(fallbackTrans);

            acElement.AppendChild(choiceElement);
            acElement.AppendChild(fallbackElement);

            // Remove existing transition or AlternateContent with transition
            foreach (var existing in slide.ChildElements
                .Where(c => c.LocalName == "transition" || c.LocalName == "AlternateContent")
                .ToList())
                existing.Remove();

            // Insert after cSld (and after any existing clrMapOvr)
            var insertAfter = slide.GetFirstChild<ColorMapOverride>() as OpenXmlElement
                ?? slide.CommonSlideData as OpenXmlElement;
            if (insertAfter != null)
                insertAfter.InsertAfterSelf(acElement);
            else
                slide.AppendChild(acElement);

            // Declare namespaces and mc:Ignorable on slide root
            // mc:Ignorable="p159" tells PowerPoint to process p159 via AlternateContent
            try { slide.AddNamespaceDeclaration("p159", p159Ns); } catch { }
            try { slide.AddNamespaceDeclaration("mc", mcNs); } catch { }
            var ignorable = slide.MCAttributes?.Ignorable?.Value;
            if (ignorable == null || !ignorable.Contains("p159"))
            {
                slide.MCAttributes ??= new MarkupCompatibilityAttributes();
                slide.MCAttributes.Ignorable = string.IsNullOrEmpty(ignorable) ? "p159" : $"{ignorable} p159";
            }

            slide.Save();
            return;
        }

        if (transElem != null) trans.Append(transElem);

        // Insert transition XML directly as an unknown element to prevent SDK from dropping it.
        // The SDK's typed Slide.Transition setter appears to work (OuterXml shows the element)
        // but Save() strips it during serialization. Workaround: inject as raw XML element
        // that the SDK preserves as-is.
        var transXml = trans.OuterXml;
        // Remove any existing transition from the slide's children (including AlternateContent wrappers for morph)
        foreach (var existing in slide.ChildElements
            .Where(c => c.LocalName == "transition" || c.LocalName == "AlternateContent")
            .ToList())
            existing.Remove();
        // Parse the transition XML as a generic OpenXmlUnknownElement and insert after cSld
        var unknownTrans = new OpenXmlUnknownElement(trans.Prefix, trans.LocalName, trans.NamespaceUri);
        unknownTrans.InnerXml = trans.InnerXml;
        foreach (var attr in trans.GetAttributes()) unknownTrans.SetAttribute(attr);
        var csd = slide.CommonSlideData;
        if (csd != null)
            csd.InsertAfterSelf(unknownTrans);
        else
            slide.AppendChild(unknownTrans);
    }

    /// <summary>Remove transition from slide by rewriting the part XML.</summary>
    private static void RewriteSlideXmlWithoutTransition(SlidePart slidePart)
    {
        slidePart.Slide?.Save();
        using var stream = slidePart.GetStream(System.IO.FileMode.Open);
        string xml;
        using (var reader = new System.IO.StreamReader(stream, leaveOpen: true))
            xml = reader.ReadToEnd();
        xml = System.Text.RegularExpressions.Regex.Replace(
            xml, @"<p:transition[^>]*?(?:/>|>.*?</p:transition>)", "",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        stream.Position = 0;
        stream.SetLength(0);
        using (var writer = new System.IO.StreamWriter(stream, leaveOpen: true))
            writer.Write(xml);
    }

    private static TransitionSlideDirectionValues ParseSlideDir(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "l" or "left" => TransitionSlideDirectionValues.Left,
            "r" or "right" => TransitionSlideDirectionValues.Right,
            "u" or "up" => TransitionSlideDirectionValues.Up,
            "d" or "down" => TransitionSlideDirectionValues.Down,
            _ => throw new ArgumentException($"Invalid slide direction: '{dir}'. Valid values: left, right, up, down.")
        };

    // For EightDirectionTransitionType where Direction is StringValue
    private static string ParseSlideDirStr(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "l" or "left" => "l",
            "r" or "right" => "r",
            "u" or "up" => "u",
            "d" or "down" => "d",
            "lu" or "leftup" => "lu",
            "ru" or "rightup" => "ru",
            "ld" or "leftdown" => "ld",
            "rd" or "rightdown" => "rd",
            _ => throw new ArgumentException($"Invalid direction: '{dir}'. Valid values: left, right, up, down, leftup, rightup, leftdown, rightdown.")
        };

    private static TransitionInOutDirectionValues ParseInOutDir(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "in" => TransitionInOutDirectionValues.In,
            "out" => TransitionInOutDirectionValues.Out,
            _ => throw new ArgumentException($"Invalid in/out direction: '{dir}'. Valid values: in, out.")
        };

    private static EnumValue<DirectionValues> ParseOrientation(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "h" or "horiz" or "horizontal" => DirectionValues.Horizontal,
            "v" or "vert" or "vertical" => DirectionValues.Vertical,
            _ => throw new ArgumentException($"Invalid orientation: '{dir}'. Valid values: horizontal, vertical.")
        };

    private static TransitionCornerDirectionValues ParseCornerDir(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "lu" or "leftup" or "upleft" => TransitionCornerDirectionValues.LeftUp,
            "ru" or "rightup" or "upright" => TransitionCornerDirectionValues.RightUp,
            "ld" or "leftdown" or "downleft" => TransitionCornerDirectionValues.LeftDown,
            "rd" or "rightdown" or "downright" => TransitionCornerDirectionValues.RightDown,
            _ => throw new ArgumentException($"Invalid corner direction: '{dir}'. Valid values: leftup, rightup, leftdown, rightdown.")
        };

    private static SplitTransition BuildSplitTransition(string? direction)
    {
        var orient = DirectionValues.Horizontal;
        var inOut = TransitionInOutDirectionValues.In;
        if (direction != null)
        {
            foreach (var token in direction.Split('-', ' '))
            {
                var t = token.ToLowerInvariant();
                if (t is "v" or "vert" or "vertical")
                    orient = DirectionValues.Vertical;
                else if (t is "h" or "horz" or "horizontal")
                    orient = DirectionValues.Horizontal;
                else if (t is "out")
                    inOut = TransitionInOutDirectionValues.Out;
                else if (t is "in")
                    inOut = TransitionInOutDirectionValues.In;
            }
        }
        return new SplitTransition { Orientation = orient, Direction = inOut };
    }

    // ==================== Shape Animations ====================

    /// <summary>
    /// Add (or remove) an entrance/exit/emphasis animation on a shape.
    /// Format: "EFFECT[-CLASS[-DURATION[-TRIGGER]]]" or "none"
    ///   EFFECT: appear, fade, fly, zoom, wipe, bounce, float, split, wheel,
    ///           spin, grow, swivel, checkerboard, blinds, bars, box, circle,
    ///           diamond, dissolve, flash, plus, random, strips, wedge
    ///   CLASS:  entrance/in/entr (default) | exit/out | emphasis/emph
    ///   DURATION: ms (default 500)
    ///   TRIGGER: click | after|afterprevious | with|withprevious
    ///            Default: first animation on slide = click, subsequent = after (sequential)
    /// Examples: "fade", "fly-entrance", "zoom-exit-800", "fade-in-500-after",
    ///           "wipe-entrance-1000-with", "fade-entrance-500-click", "none"
    /// </summary>
    private static void ApplyShapeAnimation(SlidePart slidePart, Shape shape, string value)
    {
        var slide = slidePart.Slide ?? throw new InvalidOperationException("Corrupt file");
        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
            ?? throw new ArgumentException("Shape has no ID");

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            RemoveShapeAnimations(slide, shapeId);
            return;
        }

        var parts = value.Split('-');
        var effectName = parts[0].ToLowerInvariant();

        // Flexible parsing: each segment after effect name is identified by content type
        var presetClass = TimeNodePresetClassValues.Entrance;
        var durationMs = 400;
        string? direction = null;
        AnimTrigger? explicitTrigger = null;
        int delayMs = 0, easingAccel = 0, easingDecel = 0;
        var unrecognized = new List<string>();

        for (int i = 1; i < parts.Length; i++)
        {
            var seg = parts[i].ToLowerInvariant();
            // Class?
            if (seg is "entrance" or "in" or "entr")
                presetClass = TimeNodePresetClassValues.Entrance;
            else if (seg is "exit" or "out")
                presetClass = TimeNodePresetClassValues.Exit;
            else if (seg is "emphasis" or "emph")
                presetClass = TimeNodePresetClassValues.Emphasis;
            // Trigger?
            else if (seg is "after" or "afterprevious" or "afterprev")
                explicitTrigger = AnimTrigger.AfterPrevious;
            else if (seg is "with" or "withprevious" or "withprev")
                explicitTrigger = AnimTrigger.WithPrevious;
            else if (seg is "click" or "onclick")
                explicitTrigger = AnimTrigger.OnClick;
            // Direction?
            else if (seg is "left" or "l" or "right" or "r" or "up" or "top" or "u"
                     or "down" or "bottom" or "d")
                direction = seg;
            // key=value (delay, easing, easein, easeout)?
            else if (seg.Contains('='))
            {
                var eqIdx = seg.IndexOf('=');
                var kKey = seg[..eqIdx];
                if (int.TryParse(seg[(eqIdx + 1)..], out var kVal))
                {
                    switch (kKey)
                    {
                        case "delay": delayMs = kVal; break;
                        case "easing": easingAccel = easingDecel = kVal * 1000; break;
                        case "easein": easingAccel = kVal * 1000; break;
                        case "easeout": easingDecel = kVal * 1000; break;
                    }
                }
            }
            // Duration (integer)?
            else if (int.TryParse(seg, out var d))
                durationMs = Math.Max(0, d);
            else
                unrecognized.Add(seg);
        }

        if (unrecognized.Count > 0)
            Console.Error.WriteLine($"Warning: unrecognized animation segments: {string.Join(", ", unrecognized)}. "
                + "Format: EFFECT[-CLASS][-DIRECTION][-DURATION][-TRIGGER][-delay=N][-easein=N][-easeout=N] "
                + "e.g. fly-entrance-left-400-after");

        // Resolve trigger
        AnimTrigger trigger;
        if (explicitTrigger.HasValue)
        {
            trigger = explicitTrigger.Value;
        }
        else
        {
            // Auto: first animation on slide → click, subsequent → after previous (sequential)
            var hasExistingAnimations = slide.GetFirstChild<Timing>()
                ?.Descendants<CommonTimeNode>()
                .Any(ctn => ctn.PresetId != null) ?? false;
            trigger = hasExistingAnimations ? AnimTrigger.AfterPrevious : AnimTrigger.OnClick;
        }

        // Get filter string, preset ID, and subtype from effect name
        var (presetId, filter) = GetAnimPreset(effectName, presetClass);
        var presetSubtype = GetAnimPresetSubtype(effectName, direction);
        var nodeType = trigger switch
        {
            AnimTrigger.AfterPrevious => TimeNodeValues.AfterEffect,
            AnimTrigger.WithPrevious => TimeNodeValues.WithEffect,
            _ => TimeNodeValues.ClickEffect
        };

        // Get or build timing tree
        EnsureTimingTree(slide, out var mainSeqCTn, out var bldLst);

        // Allocate IDs
        var nextId = GetMaxTimingId(slide.GetFirstChild<Timing>()!) + 1;
        var grpId = GetMaxGrpId(slide.GetFirstChild<Timing>()!);

        // The outer click-group delay depends on trigger
        var outerDelay = trigger == AnimTrigger.OnClick ? "indefinite" : "0";

        // Build the click-group par
        var clickGroup = BuildClickGroup(
            shapeId.ToString(), presetId, presetClass, nodeType,
            durationMs, filter, grpId, outerDelay, presetSubtype, ref nextId,
            delayMs, easingAccel, easingDecel);

        mainSeqCTn.ChildTimeNodeList!.AppendChild(clickGroup);

        // Update bldLst if not already there
        var shapeIdStr = shapeId.ToString();
        if (bldLst != null && !bldLst.Elements<BuildParagraph>()
                .Any(b => b.ShapeId?.Value == shapeIdStr))
        {
            bldLst.AppendChild(new BuildParagraph
            {
                ShapeId = shapeIdStr,
                GroupId = new UInt32Value((uint)grpId),
                Build = ParagraphBuildValues.AllAtOnce
            });
        }
    }

    private enum AnimTrigger { OnClick, AfterPrevious, WithPrevious }

    // ==================== Timing Helpers ====================

    private static void EnsureTimingTree(Slide slide,
        out CommonTimeNode mainSeqCTn, out BuildList? bldLst)
    {
        var timing = slide.GetFirstChild<Timing>();
        if (timing == null)
        {
            timing = new Timing();
            slide.Append(timing);
        }

        var tnLst = timing.TimeNodeList;
        if (tnLst == null)
        {
            tnLst = new TimeNodeList();
            timing.TimeNodeList = tnLst;
        }

        // Root par → cTn
        var rootPar = tnLst.GetFirstChild<ParallelTimeNode>();
        if (rootPar == null)
        {
            rootPar = new ParallelTimeNode();
            tnLst.AppendChild(rootPar);
        }

        var rootCTn = rootPar.CommonTimeNode;
        if (rootCTn == null)
        {
            rootCTn = new CommonTimeNode
            {
                Id = 1u,
                Duration = "indefinite",
                Restart = TimeNodeRestartValues.Never,
                NodeType = TimeNodeValues.TmingRoot
            };
            rootPar.CommonTimeNode = rootCTn;
        }

        var rootChildList = rootCTn.ChildTimeNodeList;
        if (rootChildList == null)
        {
            rootChildList = new ChildTimeNodeList();
            rootCTn.ChildTimeNodeList = rootChildList;
        }

        // seq element
        var seq = rootChildList.GetFirstChild<SequenceTimeNode>();
        if (seq == null)
        {
            seq = new SequenceTimeNode
            {
                Concurrent = true,
                NextAction = NextActionValues.Seek
            };
            rootChildList.AppendChild(seq);

            var seqCTn = new CommonTimeNode
            {
                Id = 2u,
                Duration = "indefinite",
                NodeType = TimeNodeValues.MainSequence
            };
            seqCTn.ChildTimeNodeList = new ChildTimeNodeList();
            seq.CommonTimeNode = seqCTn;

            // prevCondLst / nextCondLst
            var prevCondLst = new PreviousConditionList();
            prevCondLst.AppendChild(new Condition
            {
                Event = TriggerEventValues.OnPrevious,
                Delay = "0",
                TargetElement = new TargetElement(new SlideTarget())
            });
            seq.PreviousConditionList = prevCondLst;

            var nextCondLst = new NextConditionList();
            nextCondLst.AppendChild(new Condition
            {
                Event = TriggerEventValues.OnNext,
                Delay = "0",
                TargetElement = new TargetElement(new SlideTarget())
            });
            seq.NextConditionList = nextCondLst;
        }

        mainSeqCTn = seq.CommonTimeNode
            ?? throw new InvalidOperationException("seq missing cTn");
        if (mainSeqCTn.ChildTimeNodeList == null)
            mainSeqCTn.ChildTimeNodeList = new ChildTimeNodeList();

        bldLst = timing.BuildList;
        if (bldLst == null)
        {
            bldLst = new BuildList();
            timing.BuildList = bldLst;
        }
    }

    private static ParallelTimeNode BuildClickGroup(
        string shapeId,
        int presetId,
        TimeNodePresetClassValues presetClass,
        TimeNodeValues nodeType,
        int durationMs,
        string? filter,
        int grpId,
        string outerDelay,
        int presetSubtype,
        ref uint nextId,
        int delayMs = 0,
        int easingAccel = 0,
        int easingDecel = 0)
    {
        var isEntrance = presetClass == TimeNodePresetClassValues.Entrance;
        var isEmphasis = presetClass == TimeNodePresetClassValues.Emphasis;
        var animTransition = isEntrance || isEmphasis ? AnimateEffectTransitionValues.In : AnimateEffectTransitionValues.Out;

        // --- innermost cTn (the actual effect) ---
        var effectId = nextId++;
        var setVisId = nextId++;
        var animEffId = nextId++;

        var stCondEffect = new StartConditionList();
        stCondEffect.AppendChild(new Condition { Delay = "0" });

        var effectChildList = new ChildTimeNodeList();

        // p:set to make visible/hidden
        var setCTnId = nextId++;
        var setStCond = new StartConditionList();
        setStCond.AppendChild(new Condition { Delay = "0" });
        var setBehavior = new SetBehavior(
            new CommonBehavior(
                new CommonTimeNode
                {
                    Id = setVisId,
                    Duration = "1",
                    Fill = TimeNodeFillValues.Hold,
                    StartConditionList = setStCond
                },
                new TargetElement(new ShapeTarget { ShapeId = shapeId }),
                new AttributeNameList(new AttributeName("style.visibility"))
            ),
            new ToVariantValue(new StringVariantValue { Val = isEntrance || isEmphasis ? "visible" : "hidden" })
        );
        effectChildList.AppendChild(setBehavior);

        // p:animEffect
        if (filter != null)
        {
            var animEffect = new AnimateEffect
            {
                Transition = animTransition,
                Filter = filter,
                CommonBehavior = new CommonBehavior(
                    new CommonTimeNode
                    {
                        Id = animEffId,
                        Duration = durationMs.ToString()
                    },
                    new TargetElement(new ShapeTarget { ShapeId = shapeId })
                )
            };
            effectChildList.AppendChild(animEffect);
        }

        var effectCTn = new CommonTimeNode
        {
            Id = effectId,
            PresetId = presetId,
            PresetClass = presetClass,
            PresetSubtype = presetSubtype,
            Fill = TimeNodeFillValues.Hold,
            GroupId = (uint)grpId,
            NodeType = nodeType,
            StartConditionList = stCondEffect,
            ChildTimeNodeList = effectChildList
        };
        if (easingAccel > 0) effectCTn.Acceleration = easingAccel;
        if (easingDecel > 0) effectCTn.Deceleration = easingDecel;
        var effectPar = new ParallelTimeNode { CommonTimeNode = effectCTn };

        // --- middle cTn (delay wrapper) ---
        var midId = nextId++;
        var midStCond = new StartConditionList();
        midStCond.AppendChild(new Condition { Delay = delayMs > 0 ? delayMs.ToString() : "0" });
        var midChildList = new ChildTimeNodeList();
        midChildList.AppendChild(effectPar);

        var midCTn = new CommonTimeNode
        {
            Id = midId,
            Fill = TimeNodeFillValues.Hold,
            StartConditionList = midStCond,
            ChildTimeNodeList = midChildList
        };
        var midPar = new ParallelTimeNode { CommonTimeNode = midCTn };

        // --- outer click-group cTn ---
        var outerId = nextId++;
        var outerStCond = new StartConditionList();
        outerStCond.AppendChild(new Condition { Delay = outerDelay });
        var outerChildList = new ChildTimeNodeList();
        outerChildList.AppendChild(midPar);

        var outerCTn = new CommonTimeNode
        {
            Id = outerId,
            Fill = TimeNodeFillValues.Hold,
            StartConditionList = outerStCond,
            ChildTimeNodeList = outerChildList
        };
        return new ParallelTimeNode { CommonTimeNode = outerCTn };
    }

    private static void RemoveShapeAnimations(Slide slide, uint shapeId)
    {
        var timing = slide.GetFirstChild<Timing>();
        if (timing == null) return;

        var spIdStr = shapeId.ToString();

        // Remove matching ShapeTarget references deep in timing tree
        var toRemove = timing.Descendants<ShapeTarget>()
            .Where(st => st.ShapeId?.Value == spIdStr)
            .Select(st =>
            {
                // Walk up to find the top-level click-group par inside mainSeq childTnLst
                OpenXmlElement? node = st;
                while (node?.Parent != null)
                {
                    // The click-group par is a direct child of mainSeqCTn.ChildTimeNodeList
                    if (node.Parent is ChildTimeNodeList ctl &&
                        ctl.Parent is CommonTimeNode ctn &&
                        ctn.NodeType?.Value == TimeNodeValues.MainSequence)
                        return node;
                    node = node.Parent;
                }
                return null;
            })
            .Where(n => n != null)
            .Distinct()
            .ToList();

        foreach (var node in toRemove)
            node!.Remove();

        // Remove from bldLst
        var bldLst = timing.BuildList;
        if (bldLst != null)
        {
            foreach (var bp in bldLst.Elements<BuildParagraph>()
                .Where(b => b.ShapeId?.Value == shapeId.ToString()).ToList())
                bp.Remove();
        }
    }

    // ==================== Motion Path Animations ====================

    /// <summary>
    /// Apply a motion-path animation to a shape.
    /// value format: "M x y L x y E[-DURATION[-TRIGGER[-delay=N][-easing=N]]]"
    /// Coords are normalized 0.0–1.0 (relative to slide). Comma separators are normalised to spaces.
    /// Use "none" to remove existing motion path animations.
    /// </summary>
    internal static void ApplyMotionPathAnimation(SlidePart slidePart, Shape shape, string value)
    {
        var slide = slidePart.Slide ?? throw new InvalidOperationException("Corrupt file");
        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
            ?? throw new ArgumentException("Shape has no ID");

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            RemoveMotionPathAnimations(slide, shapeId);
            return;
        }

        // Split path from options at "E-" (E ends the path, options follow)
        string pathPart = value;
        int durationMs = 500;
        int delayMs = 0, easingAccel = 0, easingDecel = 0;
        var trigger = AnimTrigger.OnClick;

        var eIdx = value.IndexOf("E-", StringComparison.Ordinal);
        if (eIdx < 0) eIdx = value.IndexOf("e-", StringComparison.Ordinal);
        if (eIdx >= 0)
        {
            pathPart = value[..(eIdx + 1)]; // include the "E"
            var opts = value[(eIdx + 2)..].Split('-');
            foreach (var opt in opts)
            {
                var seg = opt.ToLowerInvariant();
                if (seg.Contains('='))
                {
                    var eq = seg.IndexOf('=');
                    if (int.TryParse(seg[(eq + 1)..], out var kVal))
                        switch (seg[..eq])
                        {
                            case "delay": delayMs = kVal; break;
                            case "easing": easingAccel = easingDecel = kVal * 1000; break;
                            case "easein": easingAccel = kVal * 1000; break;
                            case "easeout": easingDecel = kVal * 1000; break;
                        }
                }
                else if (seg is "after" or "afterprevious" or "afterprev")
                    trigger = AnimTrigger.AfterPrevious;
                else if (seg is "with" or "withprevious" or "withprev")
                    trigger = AnimTrigger.WithPrevious;
                else if (seg is "click" or "onclick")
                    trigger = AnimTrigger.OnClick;
                else if (int.TryParse(seg, out var d) && d > 0)
                    durationMs = d;
            }
        }

        pathPart = NormaliseMotionPath(pathPart);

        RemoveMotionPathAnimations(slide, shapeId);
        EnsureTimingTree(slide, out var mainSeqCTn, out _);
        var timing = slide.GetFirstChild<Timing>()!;
        var nextId = GetMaxTimingId(timing) + 1;
        var grpId = GetMaxGrpId(timing);

        var nodeType = trigger switch
        {
            AnimTrigger.AfterPrevious => TimeNodeValues.AfterEffect,
            AnimTrigger.WithPrevious => TimeNodeValues.WithEffect,
            _ => TimeNodeValues.ClickEffect
        };
        var outerDelay = trigger == AnimTrigger.OnClick ? "indefinite" : "0";

        var motionGroup = BuildMotionPathGroup(
            shapeId.ToString(), durationMs, nodeType, grpId, outerDelay,
            pathPart, ref nextId, delayMs, easingAccel, easingDecel);
        mainSeqCTn.ChildTimeNodeList!.AppendChild(motionGroup);
    }

    private static string NormaliseMotionPath(string path)
    {
        // "M0,0 L0.5,-0.3 E" → "M 0 0 L 0.5 -0.3 E"
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < path.Length; i++)
        {
            char c = path[i];
            if (char.IsLetter(c) && i > 0 && path[i - 1] != ' ')
                sb.Append(' ');
            sb.Append(c == ',' ? ' ' : c);
            if (char.IsLetter(c) && i + 1 < path.Length && path[i + 1] != ' ')
                sb.Append(' ');
        }
        // Collapse multiple spaces
        return System.Text.RegularExpressions.Regex.Replace(sb.ToString().Trim(), @" {2,}", " ");
    }

    private static ParallelTimeNode BuildMotionPathGroup(
        string shapeId, int durationMs, TimeNodeValues nodeType,
        int grpId, string outerDelay, string path,
        ref uint nextId,
        int delayMs = 0, int easingAccel = 0, int easingDecel = 0)
    {
        var effectId = nextId++;
        var animMotionId = nextId++;

        var stCond = new StartConditionList();
        stCond.AppendChild(new Condition { Delay = "0" });

        var animMotion = new AnimateMotion
        {
            Origin = AnimateMotionBehaviorOriginValues.Layout,
            PathEditMode = AnimateMotionPathEditModeValues.Relative,
            Path = path,
            CommonBehavior = new CommonBehavior(
                new CommonTimeNode { Id = animMotionId, Duration = durationMs.ToString() },
                new TargetElement(new ShapeTarget { ShapeId = shapeId })
            )
        };

        var effectCTn = new CommonTimeNode
        {
            Id = effectId,
            PresetId = 26,
            PresetSubtype = 0,
            Fill = TimeNodeFillValues.Hold,
            GroupId = (uint)grpId,
            NodeType = nodeType,
            StartConditionList = stCond,
            ChildTimeNodeList = new ChildTimeNodeList(animMotion)
        };
        effectCTn.SetAttribute(new OpenXmlAttribute("presetClass", string.Empty, "motion"));
        if (easingAccel > 0) effectCTn.Acceleration = easingAccel;
        if (easingDecel > 0) effectCTn.Deceleration = easingDecel;

        var midId = nextId++;
        var midStCond = new StartConditionList();
        midStCond.AppendChild(new Condition { Delay = delayMs > 0 ? delayMs.ToString() : "0" });
        var midCTn = new CommonTimeNode
        {
            Id = midId, Fill = TimeNodeFillValues.Hold,
            StartConditionList = midStCond,
            ChildTimeNodeList = new ChildTimeNodeList(new ParallelTimeNode { CommonTimeNode = effectCTn })
        };

        var outerId = nextId++;
        var outerStCond = new StartConditionList();
        outerStCond.AppendChild(new Condition { Delay = outerDelay });
        var outerCTn = new CommonTimeNode
        {
            Id = outerId, Fill = TimeNodeFillValues.Hold,
            StartConditionList = outerStCond,
            ChildTimeNodeList = new ChildTimeNodeList(new ParallelTimeNode { CommonTimeNode = midCTn })
        };
        return new ParallelTimeNode { CommonTimeNode = outerCTn };
    }

    private static void RemoveMotionPathAnimations(Slide slide, uint shapeId)
    {
        var timing = slide.GetFirstChild<Timing>();
        if (timing == null) return;

        var spIdStr = shapeId.ToString();
        var toRemove = timing.Descendants<ShapeTarget>()
            .Where(st => st.ShapeId?.Value == spIdStr)
            .Select(st =>
            {
                OpenXmlElement? node = st;
                while (node?.Parent != null)
                {
                    if (node.Parent is ChildTimeNodeList ctl &&
                        ctl.Parent is CommonTimeNode ctn &&
                        ctn.NodeType?.Value == TimeNodeValues.MainSequence)
                        return node;
                    node = node.Parent;
                }
                return null;
            })
            .Where(n => n != null)
            // Only remove groups that contain a motion presetClass
            .Where(n => n!.Descendants<CommonTimeNode>()
                .Any(c => c.GetAttributes().Any(a => a.LocalName == "presetClass" && a.Value == "motion")))
            .Distinct().ToList();

        foreach (var n in toRemove) n!.Remove();
    }

    private static uint GetMaxTimingId(Timing timing)
    {
        uint max = 1;
        foreach (var ctn in timing.Descendants<CommonTimeNode>())
            if (ctn.Id?.Value > max) max = ctn.Id.Value;
        return max;
    }

    private static int GetMaxGrpId(Timing timing)
    {
        int max = -1;
        foreach (var ctn in timing.Descendants<CommonTimeNode>())
        {
            var gid = (int?)ctn.GroupId?.Value;
            if (gid.HasValue && gid.Value > max) max = gid.Value;
        }
        return max + 1;
    }

    // ==================== Effect Presets ====================

    // ==================== Read back ====================

    /// <summary>
    /// Populate Format["animation"] on a shape DocumentNode by inspecting the slide Timing tree.
    /// Returns a string of the form "effectName-class-durationMs".
    /// </summary>
    private static void ReadShapeAnimation(SlidePart slidePart, Shape shape, OfficeCli.Core.DocumentNode node)
    {
        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        if (shapeId == null) return;

        var timing = slidePart.Slide?.GetFirstChild<Timing>();
        if (timing == null) return;

        var shapeIdStr = shapeId.Value.ToString();
        var shapeTarget = timing.Descendants<ShapeTarget>()
            .FirstOrDefault(st => st.ShapeId?.Value == shapeIdStr);
        if (shapeTarget == null) return;

        // Find the effect CommonTimeNode (the one with PresetClass + PresetId)
        // Skip motion path CTns (presetClass="motion" — not a valid SDK enum)
        CommonTimeNode? effectCTn = null;
        OpenXmlElement? cur = shapeTarget;
        while (cur != null)
        {
            if (cur is CommonTimeNode ctn && ctn.PresetClass != null && ctn.PresetId != null)
            {
                var rawCls = ctn.GetAttributes().FirstOrDefault(a => a.LocalName == "presetClass").Value ?? "";
                if (rawCls != "motion") { effectCTn = ctn; break; }
            }
            cur = cur.Parent;
        }
        if (effectCTn == null) return;

        var rawPresetClass = effectCTn.GetAttributes().FirstOrDefault(a => a.LocalName == "presetClass").Value ?? "";
        var cls = rawPresetClass == "exit" ? "exit"
                : rawPresetClass == "emphasis" ? "emphasis"
                : "entrance";

        // Duration from the animEffect child
        var animEffect = effectCTn.Descendants<AnimateEffect>().FirstOrDefault();
        var dur = 500;
        if (int.TryParse(animEffect?.CommonBehavior?.CommonTimeNode?.Duration, out var d)) dur = d;

        // Effect name from filter string
        var filter = animEffect?.Filter?.Value ?? "";
        var presetId = effectCTn.PresetId?.Value ?? 0;
        var effectName = filter switch
        {
            "fly"                                       => "fly",
            var f when f.StartsWith("blinds")           => "blinds",
            "box"                                       => "box",
            var f when f.StartsWith("checkerboard")     => "checkerboard",
            "circle"                                    => "circle",
            var f when f.StartsWith("crawl")            => "crawl",
            "diamond"                                   => "diamond",
            "dissolve"                                  => "dissolve",
            "fade"                                      => "fade",
            "flash"                                     => "flash",
            "plus"                                      => "plus",
            "random"                                    => "random",
            var f when f.StartsWith("barn")             => "split",
            var f when f.StartsWith("strips")           => "strips",
            "swivel"                                    => "swivel",
            "wedge"                                     => "wedge",
            var f when f.StartsWith("wheel")            => "wheel",
            var f when f.StartsWith("wipe")             => "wipe",
            "zoom"                                      => "zoom",
            "" when presetId == 1                       => "appear",
            "" when presetId == 24                      => "bounce",
            _                                           => presetId switch
            {
                1  => "appear",
                10 => "fade",
                21 => "zoom",
                24 => "bounce",
                _  => "unknown"
            }
        };

        node.Format["animation"] = $"{effectName}-{cls}-{dur}";
    }

    /// <summary>
    /// Populate Format["transition"], Format["advanceTime"], Format["advanceClick"]
    /// on a slide DocumentNode.
    /// </summary>
    /// <summary>
    /// Overload that reads transition from the SlidePart stream directly,
    /// bypassing the SDK's typed Transition accessor which may fail.
    /// </summary>
    internal static void ReadSlideTransition(SlidePart slidePart, OfficeCli.Core.DocumentNode node)
    {
        // First try SDK typed access
        var slide = slidePart.Slide;
        var trans = slide?.Transition;
        if (trans != null)
        {
            ReadSlideTransition(slide!, node);
            return;
        }

        // SDK typed access failed — try parsing from the slide's serialized XML.
        // The OuterXml may contain the transition even when the typed property is null.
        if (slide != null)
            ParseTransitionFromXml(slide.OuterXml, node);
    }

    private static void ParseTransitionFromXml(string xml, OfficeCli.Core.DocumentNode node)
    {
        var typeMatch = System.Text.RegularExpressions.Regex.Match(
            xml, @"<p:transition([^>]*?)(?:/>|>(.*?)</p:transition>)",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!typeMatch.Success) return;

        var attrs = typeMatch.Groups[1].Value;
        var inner = typeMatch.Groups[2].Value;

        // Extract transition type from first child element: <p:fade/> → "fade"
        var childMatch = System.Text.RegularExpressions.Regex.Match(inner, @"<p:(\w+)[\s/>]");
        if (childMatch.Success)
        {
            var typeName = childMatch.Groups[1].Value.ToLowerInvariant();
            node.Format["transition"] = typeName == "randombar" ? "bars" : typeName;
        }

        // Extract speed attribute
        var spdMatch = System.Text.RegularExpressions.Regex.Match(attrs, @"spd=""(\w+)""");
        if (spdMatch.Success) node.Format["transitionSpeed"] = spdMatch.Groups[1].Value;

        // Extract advance time
        var advMatch = System.Text.RegularExpressions.Regex.Match(attrs, @"advTm=""(\d+)""");
        if (advMatch.Success) node.Format["advanceTime"] = advMatch.Groups[1].Value;

        // Extract advance on click
        var clickMatch = System.Text.RegularExpressions.Regex.Match(attrs, @"advClick=""(\d+)""");
        if (clickMatch.Success) node.Format["advanceClick"] = clickMatch.Groups[1].Value == "1";
    }

    internal static void ReadSlideTransition(Slide slide, OfficeCli.Core.DocumentNode node)
    {
        var trans = slide.Transition;
        if (trans == null) return;

        // Determine type from first child element
        var transElem = trans.ChildElements.FirstOrDefault(c => c.LocalName != "extLst");
        if (transElem != null)
        {
            var typeName = transElem.LocalName.ToLowerInvariant() switch
            {
                "fade"      => "fade",
                "cut"       => "cut",
                "dissolve"  => "dissolve",
                "wipe"      => "wipe",
                "push"      => "push",
                "cover"     => "cover",
                "pull"      => "pull",
                "wheel"     => "wheel",
                "zoom"      => "zoom",
                "box"       => "zoom",
                "split"     => "split",
                "blinds"    => "blinds",
                "checker"   => "checker",
                "randombar" => "bars",
                "comb"      => "comb",
                "strips"    => "strips",
                "circle"    => "circle",
                "diamond"   => "diamond",
                "newsflash" => "newsflash",
                "plus"      => "plus",
                "random"    => "random",
                "wedge"     => "wedge",
                "flash"     => "flash",
                _           => transElem.LocalName.ToLowerInvariant()
            };
            node.Format["transition"] = typeName;
        }

        if (trans.AdvanceAfterTime != null)
            node.Format["advanceTime"] = trans.AdvanceAfterTime.Value;
        if (trans.AdvanceOnClick?.Value == false)
            node.Format["advanceClick"] = false;
    }

    /// <summary>Returns a preset subtype for the given effect name, or 0 for default.</summary>
    /// <summary>
    /// Map direction keyword to OOXML subtype. If direction is null, use effect-specific default.
    /// Subtypes: 0=none, 1=from-left, 2=from-top, 4=from-bottom, 8=from-right
    /// </summary>
    private static int GetAnimPresetSubtype(string effect, string? direction)
    {
        // If direction is explicitly specified, map it
        if (direction != null)
        {
            return direction switch
            {
                "left" or "l"                  => 8,  // object enters from left → subtype 8
                "right" or "r"                 => 2,  // from right → subtype 2
                "up" or "top" or "u"           => 1,  // from top → subtype 1
                "down" or "bottom" or "d"      => 4,  // from bottom → subtype 4
                _ => 0
            };
        }

        // Effect-specific defaults
        return effect switch
        {
            "fly" or "flyin" or "flyout" => 4,  // from bottom
            "wipe"                       => 1,   // from left
            "blinds"                     => 10,  // horizontal
            "checkerboard" or "checker"  => 5,   // across
            "strips"                     => 7,   // down-left
            "split"                      => 10,  // horizontal in
            "wheel"                      => 1,   // 1 spoke
            _                            => 0    // default
        };
    }

    /// <summary>Returns (presetId, animFilter) for the given effect name.</summary>
    private static (int presetId, string? filter) GetAnimPreset(
        string effect, TimeNodePresetClassValues cls)
    {
        if (cls == TimeNodePresetClassValues.Entrance || cls == TimeNodePresetClassValues.Exit)
        {
            return effect switch
            {
                "appear"                          => (1,  null),
                "fly" or "flyin" or "flyout"      => (2,  "fly"),
                "blinds"                          => (3,  "blinds(horizontal)"),
                "box"                             => (4,  "box"),
                "checkerboard" or "checker"       => (5,  "checkerboard(across)"),
                "circle"                          => (6,  "circle"),
                "crawlin" or "crawlout" or "crawl"=> (7,  "crawl"),
                "diamond"                         => (8,  "diamond"),
                "dissolve"                        => (9,  "dissolve"),
                "fade"                            => (10, "fade"),
                "flash" or "flashonce"            => (11, "flash"),
                "float"                           => (12, "fly"),
                "plus"                            => (13, "plus"),
                "random"                          => (14, "random"),
                "split"                           => (15, "barn(inHorizontal)"),
                "strips"                          => (16, "strips(downLeft)"),
                "swivel"                          => (17, "swivel"),
                "wedge"                           => (18, "wedge"),
                "wheel"                           => (19, "wheel(1)"),
                "wipe"                            => (20, "wipe(left)"),
                "zoom"                            => (21, "zoom"),
                "bounce"                          => (24, null),
                "swipe" or "sweep"                => (2,  "fly"),
                _ => throw new ArgumentException(
                    $"Unknown animation effect: '{effect}'. " +
                    "Supported entrance/exit effects: appear, fade, fly, zoom, wipe, bounce, float, split, " +
                    "wheel, swivel, checkerboard, blinds, dissolve, flash, box, circle, diamond, plus, strips, wedge, random. " +
                    "Use 'none' to remove.")
            };
        }
        // Emphasis
        return effect switch
        {
            "spin" or "rotate"      => (27, null),
            "grow" or "shrink"      => (26, null),
            "bold" or "boldflash"   => (1,  null),
            "wave"                  => (14, null),
            "fade"                  => (10, "fade"),
            _ => throw new ArgumentException(
                $"Unknown emphasis effect: '{effect}'. Supported: spin, grow, bold, wave, fade.")
        };
    }

    // ==================== Media Timing ====================

    /// <summary>
    /// Add a video/audio timing node to the slide's timing tree.
    /// This makes the media playable in PowerPoint (click or auto-play).
    ///
    /// Two nodes are required:
    /// 1. p:video/p:audio — media player node (in root childTnLst)
    /// 2. p:cmd cmd="playFrom(0)" — playback trigger (in main sequence, for autoplay)
    /// </summary>
    private static void AddMediaTimingNode(Slide slide, uint shapeId, bool isVideo, int volume, bool autoPlay)
    {
        EnsureTimingTree(slide, out var mainSeqCTn, out _);
        var timing = slide.GetFirstChild<Timing>()!;
        var rootCTn = timing.TimeNodeList!.GetFirstChild<ParallelTimeNode>()!.CommonTimeNode!;
        var rootChildList = rootCTn.ChildTimeNodeList!;

        var nextId = GetMaxTimingId(timing) + 1;

        // 1. Add playback command in the main sequence (triggers actual playback)
        var cmdCTn = new CommonTimeNode
        {
            Id = nextId++,
            PresetId = 1,
            PresetClass = TimeNodePresetClassValues.MediaCall,
            PresetSubtype = 0,
            Fill = TimeNodeFillValues.Hold,
            NodeType = autoPlay ? TimeNodeValues.AfterEffect : TimeNodeValues.ClickEffect
        };
        cmdCTn.StartConditionList = new StartConditionList(
            new Condition { Delay = "0" }
        );
        cmdCTn.ChildTimeNodeList = new ChildTimeNodeList(
            new Command
            {
                Type = CommandValues.Call,
                CommandName = "playFrom(0)",
                CommonBehavior = new CommonBehavior(
                    new CommonTimeNode { Id = nextId++, Duration = "1", Fill = TimeNodeFillValues.Hold },
                    new TargetElement(new ShapeTarget { ShapeId = shapeId.ToString() })
                )
            }
        );

        // Wrap in par → par → par structure for main sequence
        var innerPar = new ParallelTimeNode(new CommonTimeNode(
            new StartConditionList(new Condition { Delay = "0" }),
            new ChildTimeNodeList(new ParallelTimeNode(cmdCTn))
        ) { Id = nextId++, Fill = TimeNodeFillValues.Hold });

        var seqEntryPar = new ParallelTimeNode(new CommonTimeNode(
            new StartConditionList(new Condition { Delay = autoPlay ? "0" : "indefinite" }),
            new ChildTimeNodeList(innerPar)
        ) { Id = nextId++, Fill = TimeNodeFillValues.Hold });

        mainSeqCTn.ChildTimeNodeList ??= new ChildTimeNodeList();
        mainSeqCTn.ChildTimeNodeList.AppendChild(seqEntryPar);

        // 2. Add media player node (in root childTnLst, controls the player itself)
        var cMediaNode = new CommonMediaNode { Volume = volume };
        var mediaCTn = new CommonTimeNode
        {
            Id = nextId++,
            Fill = TimeNodeFillValues.Hold,
            Display = false
        };
        mediaCTn.StartConditionList = new StartConditionList(
            new Condition { Delay = "indefinite" }
        );
        cMediaNode.CommonTimeNode = mediaCTn;
        cMediaNode.TargetElement = new TargetElement(
            new ShapeTarget { ShapeId = shapeId.ToString() }
        );

        OpenXmlElement mediaNode;
        if (isVideo)
            mediaNode = new Video(cMediaNode) { FullScreen = false };
        else
            mediaNode = new Audio(cMediaNode) { IsNarration = false };

        rootChildList.AppendChild(mediaNode);
    }

    /// <summary>
    /// Auto-add "!!" prefix to all named shapes on the current slide and the previous slide.
    /// This ensures morph matches shapes even when their text content differs.
    /// Skips shapes that already have "!!" prefix or have default names like "TextBox N".
    /// </summary>
    private void AutoPrefixMorphNames(DocumentFormat.OpenXml.Packaging.SlidePart currentSlidePart)
    {
        var slideParts = GetSlideParts().ToList();
        var currentIdx = slideParts.IndexOf(currentSlidePart);
        if (currentIdx < 0) return;

        // Process current slide + previous slide
        // Morph on slide N means transition from slide N-1 → slide N
        var slidesToProcess = new List<DocumentFormat.OpenXml.Packaging.SlidePart> { currentSlidePart };
        if (currentIdx > 0) slidesToProcess.Add(slideParts[currentIdx - 1]);

        foreach (var sp in slidesToProcess)
        {
            var shapeTree = GetSlide(sp).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            foreach (var shape in shapeTree.Elements<Shape>())
            {
                var nvPr = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                if (nvPr?.Name == null) continue;

                var name = nvPr.Name.Value;
                if (string.IsNullOrEmpty(name)) continue;
                if (name.StartsWith("!!")) continue; // already prefixed
                // Skip auto-generated default names (TextBox N, etc.)
                if (name.StartsWith("TextBox ") || name.StartsWith("Content ") || name == "") continue;

                nvPr.Name = "!!" + name;
            }

            GetSlide(sp).Save();
        }
    }

    /// <summary>
    /// Remove "!!" prefix from shape names when morph is removed.
    /// Only strips prefix from current slide + previous slide.
    /// </summary>
    private void AutoUnprefixMorphNames(DocumentFormat.OpenXml.Packaging.SlidePart currentSlidePart)
    {
        var slideParts = GetSlideParts().ToList();
        var currentIdx = slideParts.IndexOf(currentSlidePart);
        if (currentIdx < 0) return;

        var slidesToProcess = new List<DocumentFormat.OpenXml.Packaging.SlidePart> { currentSlidePart };
        if (currentIdx > 0) slidesToProcess.Add(slideParts[currentIdx - 1]);

        foreach (var sp in slidesToProcess)
        {
            // Don't strip if this slide still has morph from another direction
            // (e.g. slide 3 has morph, user removes morph from slide 2 — don't strip slide 2
            //  if slide 3 still morphs from slide 2)
            var nextIdx = slideParts.IndexOf(sp) + 1;
            if (nextIdx < slideParts.Count)
            {
                var nextSlide = GetSlide(slideParts[nextIdx]);
                var hasMorphNext = nextSlide.ChildElements.Any(c =>
                    c.LocalName == "AlternateContent" && c.InnerXml.Contains("morph"));
                if (hasMorphNext) continue;
            }

            var shapeTree = GetSlide(sp).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            foreach (var shape in shapeTree.Elements<Shape>())
            {
                var nvPr = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                if (nvPr?.Name == null) continue;
                var name = nvPr.Name.Value;
                if (name != null && name.StartsWith("!!"))
                    nvPr.Name = name[2..];
            }

            GetSlide(sp).Save();
        }
    }
}
