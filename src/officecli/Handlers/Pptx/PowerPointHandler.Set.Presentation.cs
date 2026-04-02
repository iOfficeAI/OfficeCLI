// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// Try to handle presentation-level settings. Returns true if handled.
    /// </summary>
    private bool TrySetPresentationSetting(string key, string value)
    {
        switch (key)
        {
            // ==================== Presentation Attributes ====================
            case "firstslidenum" or "firstslidenumber":
            {
                var pres = _doc.PresentationPart!.Presentation;
                pres.FirstSlideNum = ParseHelpers.SafeParseInt(value, "firstSlideNum");
                pres.Save();
                return true;
            }
            case "rtl":
            {
                var pres = _doc.PresentationPart!.Presentation;
                pres.RightToLeft = IsTruthy(value);
                pres.Save();
                return true;
            }
            case "compatmode" or "compatibilitymode":
            {
                var pres = _doc.PresentationPart!.Presentation;
                pres.CompatibilityMode = IsTruthy(value);
                pres.Save();
                return true;
            }
            case "removepersonalinfoonsave" or "removepersonalinfo":
            {
                var pres = _doc.PresentationPart!.Presentation;
                pres.RemovePersonalInfoOnSave = IsTruthy(value);
                pres.Save();
                return true;
            }

            // ==================== PrintingProperties ====================
            case "print.what" or "printwhat":
            {
                var printProps = EnsurePrintingProperties();
                printProps.PrintWhat = value.ToLowerInvariant() switch
                {
                    "slides" => PrintOutputValues.Slides,
                    "handouts" or "handout" => PrintOutputValues.Handouts1,
                    "notes" => PrintOutputValues.Notes,
                    "outline" => PrintOutputValues.Outline,
                    _ => throw new ArgumentException($"Invalid print.what: '{value}'. Valid: slides, handouts, notes, outline")
                };
                SavePresentationProperties();
                return true;
            }
            case "print.colormode" or "printcolormode":
            {
                var printProps = EnsurePrintingProperties();
                printProps.ColorMode = value.ToLowerInvariant() switch
                {
                    "color" or "clr" => PrintColorModeValues.Color,
                    "grayscale" or "gray" => PrintColorModeValues.Gray,
                    "blackandwhite" or "bw" => PrintColorModeValues.BlackWhite,
                    _ => throw new ArgumentException($"Invalid print.colorMode: '{value}'. Valid: color, grayscale, blackAndWhite")
                };
                SavePresentationProperties();
                return true;
            }
            case "print.hiddenslides":
            {
                var printProps = EnsurePrintingProperties();
                printProps.HiddenSlides = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }
            case "print.scaletofitpaper":
            {
                var printProps = EnsurePrintingProperties();
                printProps.ScaleToFitPaper = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }
            case "print.frameslides":
            {
                var printProps = EnsurePrintingProperties();
                printProps.FrameSlides = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }

            // ==================== ShowProperties ====================
            case "show.loop" or "showloop":
            {
                var showProps = EnsureShowProperties();
                showProps.Loop = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }
            case "show.narration" or "shownarration":
            {
                var showProps = EnsureShowProperties();
                showProps.ShowNarration = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }
            case "show.animation" or "showanimation":
            {
                var showProps = EnsureShowProperties();
                showProps.ShowAnimation = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }
            case "show.usetimings" or "usetimings":
            {
                var showProps = EnsureShowProperties();
                showProps.UseTimings = IsTruthy(value);
                SavePresentationProperties();
                return true;
            }

            default:
                return false;
        }
    }

    // ==================== Helpers ====================

    private PresentationPropertiesPart EnsurePresentationPropertiesPart()
    {
        var presPart = _doc.PresentationPart!;
        return presPart.PresentationPropertiesPart
            ?? presPart.AddNewPart<PresentationPropertiesPart>();
    }

    private P.PresentationProperties EnsurePresentationPropertiesRoot()
    {
        var propsPart = EnsurePresentationPropertiesPart();
        propsPart.PresentationProperties ??= new P.PresentationProperties();
        return propsPart.PresentationProperties;
    }

    private PrintingProperties EnsurePrintingProperties()
    {
        var presProps = EnsurePresentationPropertiesRoot();
        var printProps = presProps.GetFirstChild<PrintingProperties>();
        if (printProps == null)
        {
            printProps = new PrintingProperties();
            // p:prnPr must precede p:showPr in schema order — insert before ShowProperties if present
            var showProps = presProps.GetFirstChild<ShowProperties>();
            if (showProps != null)
                showProps.InsertBeforeSelf(printProps);
            else
                presProps.AppendChild(printProps);
        }
        return printProps;
    }

    private ShowProperties EnsureShowProperties()
    {
        var presProps = EnsurePresentationPropertiesRoot();
        var showProps = presProps.GetFirstChild<ShowProperties>();
        if (showProps == null)
        {
            showProps = new ShowProperties();
            presProps.AppendChild(showProps);
        }
        return showProps;
    }

    private void SavePresentationProperties()
    {
        _doc.PresentationPart?.PresentationPropertiesPart?.PresentationProperties?.Save();
    }

    /// <summary>
    /// Read presentation-level settings into Format dictionary.
    /// </summary>
    private void PopulatePresentationSettings(DocumentNode node)
    {
        var pres = _doc.PresentationPart?.Presentation;
        if (pres == null) return;

        // Presentation attributes
        if (pres.FirstSlideNum?.Value != null && pres.FirstSlideNum.Value != 1)
            node.Format["firstSlideNum"] = pres.FirstSlideNum.Value;
        if (pres.RightToLeft?.Value == true)
            node.Format["rtl"] = true;
        if (pres.CompatibilityMode?.Value == true)
            node.Format["compatMode"] = true;
        if (pres.RemovePersonalInfoOnSave?.Value == true)
            node.Format["removePersonalInfo"] = true;

        // PresentationProperties
        var propsPart = _doc.PresentationPart?.PresentationPropertiesPart;
        var presProps = propsPart?.PresentationProperties;
        if (presProps == null) return;

        // PrintingProperties
        var printProps = presProps.GetFirstChild<PrintingProperties>();
        if (printProps != null)
        {
            if (printProps.PrintWhat?.Value != null) node.Format["print.what"] = printProps.PrintWhat.InnerText;
            if (printProps.ColorMode?.Value != null) node.Format["print.colorMode"] = printProps.ColorMode.InnerText;
            if (printProps.HiddenSlides?.Value == true) node.Format["print.hiddenSlides"] = true;
            if (printProps.ScaleToFitPaper?.Value == true) node.Format["print.scaleToFitPaper"] = true;
            if (printProps.FrameSlides?.Value == true) node.Format["print.frameSlides"] = true;
        }

        // ShowProperties
        var showProps = presProps.GetFirstChild<ShowProperties>();
        if (showProps != null)
        {
            if (showProps.Loop?.Value == true) node.Format["show.loop"] = true;
            if (showProps.ShowNarration?.Value == true) node.Format["show.narration"] = true;
            if (showProps.ShowAnimation?.Value != null) node.Format["show.animation"] = showProps.ShowAnimation.Value;
            if (showProps.UseTimings?.Value != null) node.Format["show.useTimings"] = showProps.UseTimings.Value;
        }
    }
}
