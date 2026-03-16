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
        slide.RemoveAllChildren<Transition>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
            return;

        var parts = value.Split('-');
        var typeName = parts[0].ToLowerInvariant();

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
            "split" => new SplitTransition
            {
                Orientation = ParseOrientation(direction ?? "horizontal"),
                Direction = ParseInOutDir("in")
            },
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
            _ => null
        };

        if (transElem != null) trans.Append(transElem);

        // Insert transition before timing element (schema order)
        var timing = slide.GetFirstChild<Timing>();
        if (timing != null)
            slide.InsertBefore(trans, timing);
        else
            slide.Append(trans);
    }

    private static TransitionSlideDirectionValues ParseSlideDir(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "l" or "left" => TransitionSlideDirectionValues.Left,
            "r" or "right" => TransitionSlideDirectionValues.Right,
            "u" or "up" => TransitionSlideDirectionValues.Up,
            "d" or "down" => TransitionSlideDirectionValues.Down,
            _ => TransitionSlideDirectionValues.Left
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
            _ => "l"
        };

    private static TransitionInOutDirectionValues ParseInOutDir(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "out" => TransitionInOutDirectionValues.Out,
            _ => TransitionInOutDirectionValues.In
        };

    private static EnumValue<DirectionValues> ParseOrientation(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "v" or "vert" or "vertical" => DirectionValues.Vertical,
            _ => DirectionValues.Horizontal
        };

    private static TransitionCornerDirectionValues ParseCornerDir(string dir) =>
        dir.ToLowerInvariant() switch
        {
            "lu" or "leftup" or "upleft" => TransitionCornerDirectionValues.LeftUp,
            "ru" or "rightup" or "upright" => TransitionCornerDirectionValues.RightUp,
            "ld" or "leftdown" or "downleft" => TransitionCornerDirectionValues.LeftDown,
            _ => TransitionCornerDirectionValues.RightDown
        };

    // ==================== Shape Animations ====================

    /// <summary>
    /// Add (or remove) an entrance/exit/emphasis animation on a shape.
    /// Format: "EFFECT[-CLASS[-DURATION[-TRIGGER]]]" or "none"
    ///   EFFECT: appear, fade, fly, zoom, wipe, bounce, float, split, wheel,
    ///           spin, grow, swivel, checkerboard, blinds, bars, box, circle,
    ///           diamond, dissolve, flash, plus, random, strips, wedge
    ///   CLASS:  entrance/in/entr (default) | exit/out | emphasis/emph
    ///   DURATION: ms (default 500)
    ///   TRIGGER: click (default) | after|afterprevious | with|withprevious
    /// Examples: "fade", "fly-entrance", "zoom-exit-800", "fade-in-500-after",
    ///           "wipe-entrance-1000-with", "none"
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

        // Parse class (entrance/exit/emphasis)
        var className = parts.Length > 1 ? parts[1].ToLowerInvariant() : "entrance";
        var presetClass = className switch
        {
            "exit" or "out" => TimeNodePresetClassValues.Exit,
            "emphasis" or "emph" => TimeNodePresetClassValues.Emphasis,
            _ => TimeNodePresetClassValues.Entrance
        };

        // Parse duration
        var durationMs = 500;
        if (parts.Length > 2 && int.TryParse(parts[2], out var d)) durationMs = d;

        // Parse trigger
        var triggerStr = parts.Length > 3 ? parts[3].ToLowerInvariant() : "click";
        var trigger = triggerStr switch
        {
            "after" or "afterprevious" or "afterprev" => AnimTrigger.AfterPrevious,
            "with" or "withprevious" or "withprev" => AnimTrigger.WithPrevious,
            _ => AnimTrigger.OnClick
        };

        // Get filter string and preset ID from effect name
        var (presetId, filter) = GetAnimPreset(effectName, presetClass);
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
            durationMs, filter, grpId, outerDelay, ref nextId);

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
        ref uint nextId)
    {
        var isEntrance = presetClass == TimeNodePresetClassValues.Entrance;
        var animTransition = isEntrance ? AnimateEffectTransitionValues.In : AnimateEffectTransitionValues.Out;

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
            new ToVariantValue(new StringVariantValue { Val = isEntrance ? "visible" : "hidden" })
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
            PresetSubtype = 0,
            Fill = TimeNodeFillValues.Hold,
            GroupId = (uint)grpId,
            NodeType = nodeType,
            StartConditionList = stCondEffect,
            ChildTimeNodeList = effectChildList
        };
        var effectPar = new ParallelTimeNode { CommonTimeNode = effectCTn };

        // --- middle cTn (delay wrapper) ---
        var midId = nextId++;
        var midStCond = new StartConditionList();
        midStCond.AppendChild(new Condition { Delay = "0" });
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
        CommonTimeNode? effectCTn = null;
        OpenXmlElement? cur = shapeTarget;
        while (cur != null)
        {
            if (cur is CommonTimeNode ctn && ctn.PresetClass != null && ctn.PresetId != null)
            { effectCTn = ctn; break; }
            cur = cur.Parent;
        }
        if (effectCTn == null) return;

        var clsVal = effectCTn.PresetClass?.Value;
        var cls = clsVal == TimeNodePresetClassValues.Exit ? "exit"
                : clsVal == TimeNodePresetClassValues.Emphasis ? "emphasis"
                : "entrance";

        // Duration from the animEffect child
        var animEffect = effectCTn.Descendants<AnimateEffect>().FirstOrDefault();
        var dur = 500;
        if (int.TryParse(animEffect?.CommonBehavior?.CommonTimeNode?.Duration, out var d)) dur = d;

        // Effect name from filter string
        var filter = animEffect?.Filter?.Value ?? "";
        var presetId = effectCTn.PresetId?.Value ?? 10;
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
            _                                           => "fade"
        };

        node.Format["animation"] = $"{effectName}-{cls}-{dur}";
    }

    /// <summary>
    /// Populate Format["transition"], Format["advanceTime"], Format["advanceClick"]
    /// on a slide DocumentNode.
    /// </summary>
    internal static void ReadSlideTransition(Slide slide, OfficeCli.Core.DocumentNode node)
    {
        var trans = slide.GetFirstChild<Transition>();
        if (trans == null) return;

        // Determine type from first child element
        var transElem = trans.ChildElements.FirstOrDefault(c => c is not OpenXmlLeafElement);
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
                "bounce"                          => (21, "zoom"),
                "swipe" or "sweep"                => (2,  "fly"),
                _                                 => (10, "fade")  // default: fade
            };
        }
        // Emphasis
        return effect switch
        {
            "spin" or "rotate"      => (27, null),
            "grow" or "shrink"      => (26, null),
            "bold" or "boldflash"   => (1,  null),
            "wave"                  => (14, null),
            _                       => (10, null)
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
}
