// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

// Per-element-type Add helpers for conditional-formatting paths (cf, databar, colorscale, iconset, formulacf, cellis, cfextended-group). Mechanically extracted from the Add() god-method.
public partial class ExcelHandler
{
    private string AddCf(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Dispatch to specific CF type based on "type" (primary) or "rule" (alias) property.
        // R2-2: `rule=cellIs` is also accepted — user expectation from real Excel vocabulary
        // (Excel calls these "rules", OOXML calls them cfRule "type").
        var cfTypeRaw = properties.GetValueOrDefault("type")
            ?? properties.GetValueOrDefault("rule");
        var cfType = (cfTypeRaw ?? "databar").ToLowerInvariant();
        return cfType switch
        {
            "databar" => Add(parentPath, "conditionalformatting", position, properties),
            "iconset" => Add(parentPath, "iconset", position, properties),
            "colorscale" => Add(parentPath, "colorscale", position, properties),
            "formula" or "expression" => Add(parentPath, "formulacf", position, properties),
            "cellis" => Add(parentPath, "cellis", position, properties),
            // R39-1: `top` / `topPercent` / `bottom` / `bottomPercent` are
            // user-facing aliases for the OOXML `top10` cfRule. Without this
            // mapping, the dispatch fell through to the default `databar`
            // branch and silently rewrote the rule type. Set `percent`/
            // `bottom` properties so the topn branch emits the right attrs.
            "topn" or "top10" or "top" => Add(parentPath, "topn", position, properties),
            "toppercent" => AddTopRouted(parentPath, position, properties, percent: true, bottom: false),
            "bottom" => AddTopRouted(parentPath, position, properties, percent: false, bottom: true),
            "bottompercent" => AddTopRouted(parentPath, position, properties, percent: true, bottom: true),
            "aboveaverage" => Add(parentPath, "aboveaverage", position, properties),
            "uniquevalues" => Add(parentPath, "uniquevalues", position, properties),
            "duplicatevalues" => Add(parentPath, "duplicatevalues", position, properties),
            "containstext" => Add(parentPath, "containstext", position, properties),
            "dateoccurring" or "timeperiod" => Add(parentPath, "dateoccurring", position, properties),
            "belowaverage" or "containsblanks" or "notcontainsblanks" or "containserrors" or "notcontainserrors" or "contains" or "notcontains" or "beginswith" or "endswith"
                => Add(parentPath, "cfextended", position, properties),
            // Reject unknown CF types instead of silently falling back to
            // dataBar — silent fallback hides typos like `type=badtype` and
            // produces a rule the user did not ask for.
            _ => throw new ArgumentException(
                $"Unknown CF type '{cfTypeRaw}'. Valid: databar, iconset, colorscale, formula, cellIs, "
                + "top10, topPercent, bottom, bottomPercent, aboveAverage, belowAverage, "
                + "uniqueValues, duplicateValues, containsText, notContains, beginsWith, endsWith, "
                + "containsBlanks, notContainsBlanks, containsErrors, notContainsErrors, dateOccurring.")
        };
    }

    // R39-1: thread `percent`/`bottom` flags into the topn branch so that
    // `type=topPercent` / `type=bottom` / `type=bottomPercent` route to the
    // same `top10` cfRule with the right attributes set, instead of falling
    // through to the dataBar default. Mutates `properties` in place; keys
    // already supplied by the user take precedence.
    private string AddTopRouted(string parentPath, InsertPosition? position, Dictionary<string, string> properties, bool percent, bool bottom)
    {
        if (!properties.ContainsKey("percent")) properties["percent"] = percent ? "true" : "false";
        if (!properties.ContainsKey("bottom")) properties["bottom"] = bottom ? "true" : "false";
        return Add(parentPath, "topn", position, properties);
    }

    private string AddDataBar(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Dispatch to specific CF type if "type" or "rule" property is specified.
        // R2-2: `rule=` is an accepted alias for `type=` (matches Excel UI vocabulary).
        var cfTypeProp = properties.GetValueOrDefault("type") ?? properties.GetValueOrDefault("rule");
        if (cfTypeProp != null)
        {
            var cfTypeLower = cfTypeProp.ToLowerInvariant();
            if (cfTypeLower is "iconset") return Add(parentPath, "iconset", position, properties);
            if (cfTypeLower is "colorscale") return Add(parentPath, "colorscale", position, properties);
            if (cfTypeLower is "formula" or "expression") return Add(parentPath, "formulacf", position, properties);
            if (cfTypeLower is "cellis") return Add(parentPath, "cellis", position, properties);
            // R39-1: same alias set as AddCf — keep both dispatch sites in sync.
            if (cfTypeLower is "topn" or "top10" or "top") return Add(parentPath, "topn", position, properties);
            if (cfTypeLower is "toppercent") return AddTopRouted(parentPath, position, properties, percent: true, bottom: false);
            if (cfTypeLower is "bottom") return AddTopRouted(parentPath, position, properties, percent: false, bottom: true);
            if (cfTypeLower is "bottompercent") return AddTopRouted(parentPath, position, properties, percent: true, bottom: true);
            if (cfTypeLower is "aboveaverage") return Add(parentPath, "aboveaverage", position, properties);
            if (cfTypeLower is "uniquevalues") return Add(parentPath, "uniquevalues", position, properties);
            if (cfTypeLower is "duplicatevalues") return Add(parentPath, "duplicatevalues", position, properties);
            if (cfTypeLower is "containstext") return Add(parentPath, "containstext", position, properties);
            if (cfTypeLower is "dateoccurring" or "timeperiod") return Add(parentPath, "dateoccurring", position, properties);
            if (cfTypeLower is "belowaverage" or "containsblanks" or "notcontainsblanks" or "containserrors" or "notcontainserrors" or "contains" or "notcontains" or "beginswith" or "endswith")
                return Add(parentPath, "cfextended", position, properties);
            // R10: Reject unknown CF types instead of silently falling through to
            // dataBar. The `cf` alias (AddCf) already throws on unknowns; mirror
            // the behavior here so both `--type cf` and `--type conditionalformatting`
            // share the same allowlist (CONSISTENCY(cf-type-allowlist)). `databar`
            // is the documented default for this alias, so allow it explicitly.
            if (cfTypeLower is not "databar")
            {
                throw new ArgumentException(
                    $"Unknown CF type '{cfTypeProp}'. Valid: databar, iconset, colorscale, formula, cellIs, "
                    + "top10, topPercent, bottom, bottomPercent, aboveAverage, belowAverage, "
                    + "uniqueValues, duplicateValues, containsText, notContains, beginsWith, endsWith, "
                    + "containsBlanks, notContainsBlanks, containsErrors, notContainsErrors, dateOccurring.");
            }
        }
        var cfSegments = parentPath.TrimStart('/').Split('/', 2);
        var cfSheetName = cfSegments[0];
        var cfWorksheet = FindWorksheet(cfSheetName)
            ?? throw new ArgumentException($"Sheet not found: {cfSheetName}");

        var sqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
        var minVal = properties.ContainsKey("min") ? properties["min"] : (string?)null;
        var maxVal = properties.ContainsKey("max") ? properties["max"] : (string?)null;
        var cfColor = properties.GetValueOrDefault("color", "638EC6");
        var normalizedColor = ParseHelpers.NormalizeArgbColor(cfColor);

        var cfRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.DataBar,
            Priority = NextCfPriority(GetSheet(cfWorksheet))
        };
        var dataBar = new DataBar();
        // R10-1: when cfvo type is min/max, omit `val` attribute (Excel rejects val="").
        var dbMinCfvo = new ConditionalFormatValueObject
        {
            Type = minVal != null ? ConditionalFormatValueObjectValues.Number : ConditionalFormatValueObjectValues.Min
        };
        if (minVal != null) dbMinCfvo.Val = minVal;
        dataBar.Append(dbMinCfvo);
        var dbMaxCfvo = new ConditionalFormatValueObject
        {
            Type = maxVal != null ? ConditionalFormatValueObjectValues.Number : ConditionalFormatValueObjectValues.Max
        };
        if (maxVal != null) dbMaxCfvo.Val = maxVal;
        dataBar.Append(dbMaxCfvo);
        dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedColor });
        cfRule.Append(dataBar);
        // CF6 — dataBar `showValue=false` hides the cell's numeric
        // value under the bar. Defaults to true in OOXML; only emit
        // the attribute when the user opted out.
        if (properties.TryGetValue("showValue", out var dbShowVal) && !ParseHelpers.IsTruthy(dbShowVal))
            dataBar.ShowValue = false;
        ApplyStopIfTrue(cfRule, properties);

        // R10-1: Also emit Excel 2010+ x14 extension so negative values
        // render leftward in red with an axis. Without this block, Excel
        // uses the 2007 dataBar which treats all values as positive
        // (rightward blue bars, no axis, no red for negatives).
        var dbGuid = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        // Attach x14:id extension onto the 2007 cfRule so it's paired
        // with the sibling x14:cfRule in the worksheet extLst.
        var dbRuleExtList = new ConditionalFormattingRuleExtensionList();
        var dbRuleExt = new ConditionalFormattingRuleExtension
        {
            Uri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
        };
        dbRuleExt.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        dbRuleExt.Append(new X14.Id(dbGuid));
        dbRuleExtList.Append(dbRuleExt);
        cfRule.Append(dbRuleExtList);

        var cf = new ConditionalFormatting(cfRule)
        {
            SequenceOfReferences = new ListValue<StringValue>(
                sqref.Split(' ').Select(s => new StringValue(s)))
        };

        var wsElement = GetSheet(cfWorksheet);
        InsertConditionalFormatting(wsElement, cf);

        // R10-1: Build the x14:dataBar counterpart under worksheet extLst.
        var dbNegColor = ParseHelpers.NormalizeArgbColor(properties.GetValueOrDefault("negativeColor", "FF0000"));
        var dbAxisColor = ParseHelpers.NormalizeArgbColor(properties.GetValueOrDefault("axisColor", "000000"));
        var dbAxisPos = (properties.GetValueOrDefault("axisPosition") ?? "automatic").ToLowerInvariant();
        var dbAxisPosVal = dbAxisPos switch
        {
            "middle" => X14.DataBarAxisPositionValues.Middle,
            "none" => X14.DataBarAxisPositionValues.None,
            _ => X14.DataBarAxisPositionValues.Automatic
        };

        var x14DataBar = new X14.DataBar
        {
            MinLength = 0U,
            MaxLength = 100U,
            AxisPosition = dbAxisPosVal
        };
        var x14MinCfvo = new X14.ConditionalFormattingValueObject
        {
            Type = minVal != null
                ? X14.ConditionalFormattingValueObjectTypeValues.Numeric
                : X14.ConditionalFormattingValueObjectTypeValues.AutoMin
        };
        if (minVal != null) x14MinCfvo.Append(new DocumentFormat.OpenXml.Office.Excel.Formula(minVal));
        x14DataBar.Append(x14MinCfvo);
        var x14MaxCfvo = new X14.ConditionalFormattingValueObject
        {
            Type = maxVal != null
                ? X14.ConditionalFormattingValueObjectTypeValues.Numeric
                : X14.ConditionalFormattingValueObjectTypeValues.AutoMax
        };
        if (maxVal != null) x14MaxCfvo.Append(new DocumentFormat.OpenXml.Office.Excel.Formula(maxVal));
        x14DataBar.Append(x14MaxCfvo);
        x14DataBar.Append(new X14.FillColor { Rgb = normalizedColor });
        x14DataBar.Append(new X14.NegativeFillColor { Rgb = dbNegColor });
        x14DataBar.Append(new X14.BarAxisColor { Rgb = dbAxisColor });

        var x14CfRule = new X14.ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.DataBar,
            Id = dbGuid
        };
        x14CfRule.Append(x14DataBar);

        var x14Cf = new X14.ConditionalFormatting();
        x14Cf.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");
        x14Cf.Append(x14CfRule);
        x14Cf.Append(new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(sqref));

        EnsureWorksheetX14ConditionalFormatting(wsElement, x14Cf);

        SaveWorksheet(cfWorksheet);
        var dbCfCount = wsElement.Elements<ConditionalFormatting>().Count();
        return $"/{cfSheetName}/cf[{dbCfCount}]";
    }

    private string AddColorScale(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var csSegments = parentPath.TrimStart('/').Split('/', 2);
        var csSheetName = csSegments[0];
        var csWorksheet = FindWorksheet(csSheetName)
            ?? throw new ArgumentException($"Sheet not found: {csSheetName}");

        // CONSISTENCY(cf-sqref): three-level fallback matches dataBar/formulacf branches
        var csSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
        var minColor = properties.GetValueOrDefault("mincolor", "F8696B");
        var maxColor = properties.GetValueOrDefault("maxcolor", "63BE7B");
        var midColor = properties.GetValueOrDefault("midcolor");

        var normalizedMinColor = ParseHelpers.NormalizeArgbColor(minColor);
        var normalizedMaxColor = ParseHelpers.NormalizeArgbColor(maxColor);

        // CF5 — accept user-supplied midpoint percentile (`midpoint=50`, default 50).
        var midPointStr = properties.GetValueOrDefault("midpoint")
            ?? properties.GetValueOrDefault("midPoint")
            ?? "50";
        var colorScale = new ColorScale();
        colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
        if (midColor != null)
            colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = midPointStr });
        colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
        colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedMinColor });
        if (midColor != null)
        {
            var normalizedMidColor = ParseHelpers.NormalizeArgbColor(midColor);
            colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedMidColor });
        }
        colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedMaxColor });

        var csRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.ColorScale,
            Priority = NextCfPriority(GetSheet(csWorksheet))
        };
        csRule.Append(colorScale);
        ApplyStopIfTrue(csRule, properties);

        var csCf = new ConditionalFormatting(csRule)
        {
            SequenceOfReferences = new ListValue<StringValue>(
                csSqref.Split(' ').Select(s => new StringValue(s)))
        };

        var csWsElement = GetSheet(csWorksheet);
        InsertConditionalFormatting(csWsElement, csCf);

        SaveWorksheet(csWorksheet);
        var csCfCount = csWsElement.Elements<ConditionalFormatting>().Count();
        return $"/{csSheetName}/cf[{csCfCount}]";
    }

    private string AddIconSet(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var isSegments = parentPath.TrimStart('/').Split('/', 2);
        var isSheetName = isSegments[0];
        var isWorksheet = FindWorksheet(isSheetName)
            ?? throw new ArgumentException($"Sheet not found: {isSheetName}");

        // CONSISTENCY(cf-sqref): three-level fallback matches dataBar/formulacf branches
        var isSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
        var iconSetName = properties.GetValueOrDefault("iconset") ?? properties.GetValueOrDefault("icons", "3TrafficLights1");
        var reverse = properties.TryGetValue("reverse", out var revVal) && IsTruthy(revVal);
        var showValue = !properties.TryGetValue("showvalue", out var svVal) || IsTruthy(svVal);

        var iconSetVal = ParseIconSetValues(iconSetName);

        var iconSet = new IconSet { IconSetValue = iconSetVal };
        if (reverse) iconSet.Reverse = true;
        if (!showValue) iconSet.ShowValue = false;

        // Add threshold values based on icon count
        var iconCount = GetIconCount(iconSetName);
        for (int i = 0; i < iconCount; i++)
        {
            if (i == 0)
                iconSet.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Percent,
                    Val = "0"
                });
            else
                iconSet.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Percent,
                    Val = (i * 100 / iconCount).ToString()
                });
        }

        var isRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.IconSet,
            Priority = NextCfPriority(GetSheet(isWorksheet))
        };
        isRule.Append(iconSet);
        ApplyStopIfTrue(isRule, properties);

        var isCf = new ConditionalFormatting(isRule)
        {
            SequenceOfReferences = new ListValue<StringValue>(
                isSqref.Split(' ').Select(s => new StringValue(s)))
        };

        var isWsElement = GetSheet(isWorksheet);
        InsertConditionalFormatting(isWsElement, isCf);

        SaveWorksheet(isWorksheet);
        var isCfCount = isWsElement.Elements<ConditionalFormatting>().Count();
        return $"/{isSheetName}/cf[{isCfCount}]";
    }

    private string AddFormulaCf(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var fcfSegments = parentPath.TrimStart('/').Split('/', 2);
        var fcfSheetName = fcfSegments[0];
        var fcfWorksheet = FindWorksheet(fcfSheetName)
            ?? throw new ArgumentException($"Sheet not found: {fcfSheetName}");

        // CONSISTENCY(cf-sqref): three-level fallback matches dataBar/colorScale branches
        var fcfSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
        var fcfFormula = properties.GetValueOrDefault("formula")
            ?? throw new ArgumentException("Formula-based conditional formatting requires 'formula' property (e.g. formula=$A1>100)");

        // Build DifferentialFormat (dxf) for the formatting.
        // A dxf Font may carry: b, i, u, strike, sz, rFont, color.
        // All sub-props are threaded together so users can combine
        // (e.g. bold + italic + underline + custom size + name).
        var dxf = new DifferentialFormat();
        var dxfFont = BuildFormulaCfFont(properties);
        if (dxfFont != null) dxf.Append(dxfFont);

        if (properties.TryGetValue("fill", out var fillColor))
        {
            var normalizedFillColor = ParseHelpers.NormalizeArgbColor(fillColor);
            dxf.Append(new Fill(new PatternFill(
                new BackgroundColor { Rgb = normalizedFillColor })
            { PatternType = PatternValues.Solid }));
        }

        // Add dxf to stylesheet (ensure it exists)
        var fcfWbPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("Workbook not found");
        var fcfStyleMgr = new ExcelStyleManager(fcfWbPart);
        fcfStyleMgr.EnsureStylesPart();
        var stylesheet = fcfWbPart.WorkbookStylesPart!.Stylesheet!;

        var dxfs = stylesheet.GetFirstChild<DifferentialFormats>();
        if (dxfs == null)
        {
            dxfs = new DifferentialFormats { Count = 0 };
            stylesheet.Append(dxfs);
        }
        dxfs.Append(dxf);
        dxfs.Count = (uint)dxfs.Elements<DifferentialFormat>().Count();
        _dirtyStylesheet = true;

        var dxfId = dxfs.Count!.Value - 1;

        var fcfRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.Expression,
            Priority = NextCfPriority(GetSheet(fcfWorksheet)),
            FormatId = dxfId
        };
        fcfRule.Append(new Formula(fcfFormula));
        ApplyStopIfTrue(fcfRule, properties);

        var fcfCf = new ConditionalFormatting(fcfRule)
        {
            SequenceOfReferences = new ListValue<StringValue>(
                fcfSqref.Split(' ').Select(s => new StringValue(s)))
        };

        var fcfWsElement = GetSheet(fcfWorksheet);
        InsertConditionalFormatting(fcfWsElement, fcfCf);

        SaveWorksheet(fcfWorksheet);
        var fcfCfCount = fcfWsElement.Elements<ConditionalFormatting>().Count();
        return $"/{fcfSheetName}/cf[{fcfCfCount}]";
    }

    private string AddCellIs(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // R2-2: cellIs conditional formatting — compare each cell value against
        // a literal (or formula) using one of greaterThan/lessThan/... operators.
        var cisSegments = parentPath.TrimStart('/').Split('/', 2);
        var cisSheetName = cisSegments[0];
        var cisWorksheet = FindWorksheet(cisSheetName)
            ?? throw new ArgumentException($"Sheet not found: {cisSheetName}");

        // CONSISTENCY(cf-sqref): three-level fallback matches dataBar/colorScale branches
        var cisSqref = properties.GetValueOrDefault("sqref")
            ?? properties.GetValueOrDefault("range")
            ?? properties.GetValueOrDefault("ref", "A1:A10");
        var opStr = (properties.GetValueOrDefault("operator") ?? "greaterThan").Trim();
        var opVal = opStr.ToLowerInvariant() switch
        {
            "greaterthan" or "gt" or ">" => ConditionalFormattingOperatorValues.GreaterThan,
            "lessthan" or "lt" or "<" => ConditionalFormattingOperatorValues.LessThan,
            "greaterthanorequal" or "gte" or ">=" => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
            "lessthanorequal" or "lte" or "<=" => ConditionalFormattingOperatorValues.LessThanOrEqual,
            "equal" or "eq" or "=" or "==" => ConditionalFormattingOperatorValues.Equal,
            "notequal" or "ne" or "!=" or "<>" => ConditionalFormattingOperatorValues.NotEqual,
            "between" => ConditionalFormattingOperatorValues.Between,
            "notbetween" => ConditionalFormattingOperatorValues.NotBetween,
            _ => throw new ArgumentException(
                $"Unsupported cellIs operator '{opStr}'. Valid: greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, equal, notEqual, between, notBetween.")
        };

        var primary = properties.GetValueOrDefault("value")
            ?? properties.GetValueOrDefault("formula")
            ?? properties.GetValueOrDefault("value1")
            ?? throw new ArgumentException("cellIs conditional formatting requires 'value' property (e.g. value=50).");
        var secondary = properties.GetValueOrDefault("value2")
            ?? properties.GetValueOrDefault("maxvalue");

        // Build DifferentialFormat (dxf)
        var cisDxf = new DifferentialFormat();
        if (properties.TryGetValue("font.color", out var cisFontColor))
        {
            var normalizedFontColor = ParseHelpers.NormalizeArgbColor(cisFontColor);
            cisDxf.Append(new Font(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedFontColor }));
        }
        if (properties.TryGetValue("font.bold", out var cisFontBold) && IsTruthy(cisFontBold))
        {
            var existingFont = cisDxf.GetFirstChild<Font>();
            if (existingFont != null) existingFont.Append(new Bold());
            else cisDxf.Append(new Font(new Bold()));
        }
        if (properties.TryGetValue("fill", out var cisFill))
        {
            var normalizedFill = ParseHelpers.NormalizeArgbColor(cisFill);
            cisDxf.Append(new Fill(new PatternFill(
                new BackgroundColor { Rgb = normalizedFill })
            { PatternType = PatternValues.Solid }));
        }

        var cisWbPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("Workbook not found");
        var cisStyleMgr = new ExcelStyleManager(cisWbPart);
        cisStyleMgr.EnsureStylesPart();
        var cisStylesheet = cisWbPart.WorkbookStylesPart!.Stylesheet!;
        var cisDxfs = cisStylesheet.GetFirstChild<DifferentialFormats>();
        if (cisDxfs == null)
        {
            cisDxfs = new DifferentialFormats { Count = 0 };
            cisStylesheet.Append(cisDxfs);
        }
        cisDxfs.Append(cisDxf);
        cisDxfs.Count = (uint)cisDxfs.Elements<DifferentialFormat>().Count();
        _dirtyStylesheet = true;
        var cisDxfId = cisDxfs.Count!.Value - 1;

        var cisRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.CellIs,
            Priority = NextCfPriority(GetSheet(cisWorksheet)),
            FormatId = cisDxfId,
            Operator = opVal
        };
        cisRule.Append(new Formula(primary));
        if ((opVal == ConditionalFormattingOperatorValues.Between
             || opVal == ConditionalFormattingOperatorValues.NotBetween)
            && secondary != null)
        {
            cisRule.Append(new Formula(secondary));
        }
        ApplyStopIfTrue(cisRule, properties);

        var cisCf = new ConditionalFormatting(cisRule)
        {
            SequenceOfReferences = new ListValue<StringValue>(
                cisSqref.Split(' ').Select(s => new StringValue(s)))
        };

        var cisWsElement = GetSheet(cisWorksheet);
        InsertConditionalFormatting(cisWsElement, cisCf);

        SaveWorksheet(cisWorksheet);
        var cisCfCount = cisWsElement.Elements<ConditionalFormatting>().Count();
        return $"/{cisSheetName}/cf[{cisCfCount}]";
    }

    private string AddCfExtended(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var cfNewSegments = parentPath.TrimStart('/').Split('/', 2);
        var cfNewSheetName = cfNewSegments[0];
        var cfNewWorksheet = FindWorksheet(cfNewSheetName)
            ?? throw new ArgumentException($"Sheet not found: {cfNewSheetName}");
        var cfNewSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
        var cfNewPriority = NextCfPriority(GetSheet(cfNewWorksheet));

        ConditionalFormattingRule cfNewRule;
        var typeLower = type.ToLowerInvariant();
        // For cfextended dispatch, the actual requested sub-type is in
        // properties["type"] (the user-facing switch; the outer `type`
        // variable is literal "cfextended" here).
        if (typeLower == "cfextended")
            typeLower = (properties.GetValueOrDefault("type", "") ?? "").ToLowerInvariant();

        switch (typeLower)
        {
            case "topn":
            {
                // Accept `rank=` (OOXML attribute name), `top=`/`bottomN=` (legacy
                // aliases), and `value=` (R26-1: matches the cellIs vocabulary so
                // users don't have to learn separate names per CF subtype).
                var rankStr = properties.GetValueOrDefault("rank")
                    ?? properties.GetValueOrDefault("top")
                    ?? properties.GetValueOrDefault("bottomN")
                    ?? properties.GetValueOrDefault("value")
                    ?? "10";
                if (!int.TryParse(rankStr, out var rankInt))
                    throw new ArgumentException(
                        $"top10 conditional formatting requires an integer rank (got '{rankStr}'). Use top=N or value=N.");
                if (rankInt <= 0)
                    throw new ArgumentException(
                        $"top10 conditional formatting requires rank >= 1 (got {rankInt}).");
                var rank = (uint)rankInt;
                var percent = ParseHelpers.IsTruthy(properties.GetValueOrDefault("percent", "false"));
                var bottom = ParseHelpers.IsTruthy(properties.GetValueOrDefault("bottom", "false"));
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.Top10,
                    Priority = cfNewPriority,
                    Rank = rank,
                    Percent = percent ? true : null,
                    Bottom = bottom ? true : null
                };
                break;
            }
            case "aboveaverage":
            {
                // `above=` is the legacy spelling; `aboveaverage=false`
                // (matching the cfType name) is accepted as an alias
                // so users can mirror the OOXML attribute.
                var aboveBelow = properties.GetValueOrDefault("above",
                    properties.GetValueOrDefault("aboveaverage", "true"));
                var aboveRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.AboveAverage,
                    Priority = cfNewPriority,
                    AboveAverage = ParseHelpers.IsTruthy(aboveBelow) ? null : false
                };
                // R15-3: wire stdDev= (deviations above/below mean)
                // and equalAverage= (include values equal to the mean)
                // onto the cfRule.
                if (properties.TryGetValue("stdDev", out var stdDevRaw)
                    && !string.IsNullOrWhiteSpace(stdDevRaw)
                    && int.TryParse(stdDevRaw, out var stdDevVal))
                {
                    aboveRule.StdDev = stdDevVal;
                }
                if (properties.TryGetValue("equalAverage", out var eqAvgRaw)
                    && !string.IsNullOrWhiteSpace(eqAvgRaw)
                    && ParseHelpers.IsTruthy(eqAvgRaw))
                {
                    aboveRule.EqualAverage = true;
                }
                cfNewRule = aboveRule;
                break;
            }
            case "uniquevalues":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.UniqueValues,
                    Priority = cfNewPriority
                };
                break;
            }
            case "duplicatevalues":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.DuplicateValues,
                    Priority = cfNewPriority
                };
                break;
            }
            case "containstext":
            {
                var text = properties.GetValueOrDefault("text", "");
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.ContainsText,
                    Priority = cfNewPriority,
                    Text = text,
                    Operator = ConditionalFormattingOperatorValues.ContainsText
                };
                var firstCell = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"NOT(ISERROR(SEARCH(\"{text}\",{firstCell})))"));
                break;
            }
            case "dateoccurring":
            {
                // Accept both `period=` (docs/canonical) and `timePeriod=`
                // (OOXML attribute spelling) as input aliases.
                var period = properties.GetValueOrDefault("period")
                    ?? properties.GetValueOrDefault("timePeriod")
                    ?? properties.GetValueOrDefault("timeperiod")
                    ?? "today";
                var normalizedPeriod = period.ToLowerInvariant() switch
                {
                    "today" => "today",
                    "yesterday" => "yesterday",
                    "tomorrow" => "tomorrow",
                    "last7days" => "last7Days",
                    "thisweek" => "thisWeek",
                    "lastweek" => "lastWeek",
                    "nextweek" => "nextWeek",
                    "thismonth" => "thisMonth",
                    "lastmonth" => "lastMonth",
                    "nextmonth" => "nextMonth",
                    _ => period
                };
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.TimePeriod,
                    Priority = cfNewPriority,
                    TimePeriod = new EnumValue<TimePeriodValues>(normalizedPeriod switch
                    {
                        "today" => TimePeriodValues.Today,
                        "yesterday" => TimePeriodValues.Yesterday,
                        "tomorrow" => TimePeriodValues.Tomorrow,
                        "last7Days" => TimePeriodValues.Last7Days,
                        "thisWeek" => TimePeriodValues.ThisWeek,
                        "lastWeek" => TimePeriodValues.LastWeek,
                        "nextWeek" => TimePeriodValues.NextWeek,
                        "thisMonth" => TimePeriodValues.ThisMonth,
                        "lastMonth" => TimePeriodValues.LastMonth,
                        "nextMonth" => TimePeriodValues.NextMonth,
                        _ => TimePeriodValues.Today
                    })
                };
                break;
            }
            case "belowaverage":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.AboveAverage,
                    Priority = cfNewPriority,
                    AboveAverage = false
                };
                break;
            }
            case "containsblanks":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.ContainsBlanks,
                    Priority = cfNewPriority
                };
                var fc0 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"LEN(TRIM({fc0}))=0"));
                break;
            }
            case "notcontainsblanks":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.NotContainsBlanks,
                    Priority = cfNewPriority
                };
                var fc1 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"LEN(TRIM({fc1}))>0"));
                break;
            }
            case "containserrors":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.ContainsErrors,
                    Priority = cfNewPriority
                };
                var fc2 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"ISERROR({fc2})"));
                break;
            }
            case "notcontainserrors":
            {
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.NotContainsErrors,
                    Priority = cfNewPriority
                };
                var fc3 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"NOT(ISERROR({fc3}))"));
                break;
            }
            case "contains":
            {
                var ctext = properties.GetValueOrDefault("text", "");
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.ContainsText,
                    Priority = cfNewPriority,
                    Text = ctext,
                    Operator = ConditionalFormattingOperatorValues.ContainsText
                };
                var fc4 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"NOT(ISERROR(SEARCH(\"{ctext}\",{fc4})))"));
                break;
            }
            case "notcontains":
            {
                var nctext = properties.GetValueOrDefault("text", "");
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.NotContainsText,
                    Priority = cfNewPriority,
                    Text = nctext,
                    Operator = ConditionalFormattingOperatorValues.NotContains
                };
                var fc5 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"ISERROR(SEARCH(\"{nctext}\",{fc5}))"));
                break;
            }
            case "beginswith":
            {
                var btext = properties.GetValueOrDefault("text", "");
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.BeginsWith,
                    Priority = cfNewPriority,
                    Text = btext,
                    Operator = ConditionalFormattingOperatorValues.BeginsWith
                };
                var fc6 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"LEFT({fc6},{btext.Length})=\"{btext}\""));
                break;
            }
            case "endswith":
            {
                var etext = properties.GetValueOrDefault("text", "");
                cfNewRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.EndsWith,
                    Priority = cfNewPriority,
                    Text = etext,
                    Operator = ConditionalFormattingOperatorValues.EndsWith
                };
                var fc7 = cfNewSqref.Split(':')[0].TrimStart('$');
                cfNewRule.AppendChild(new Formula($"RIGHT({fc7},{etext.Length})=\"{etext}\""));
                break;
            }
            default:
                throw new ArgumentException($"Unsupported CF type: {typeLower}");
        }

        ApplyStopIfTrue(cfNewRule, properties);

        // Build DXF formatting if fill/font properties are provided
        var cfNewDxf = new DifferentialFormat();
        bool cfNewHasDxf = false;
        if (properties.TryGetValue("font.color", out var cfNewFontColor))
        {
            var normalizedFontColor = ParseHelpers.NormalizeArgbColor(cfNewFontColor);
            cfNewDxf.Append(new Font(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedFontColor }));
            cfNewHasDxf = true;
        }
        else if (properties.TryGetValue("font.bold", out var cfNewFontBold) && IsTruthy(cfNewFontBold))
        {
            cfNewDxf.Append(new Font(new Bold()));
            cfNewHasDxf = true;
        }
        if (properties.TryGetValue("fill", out var cfNewFillColor))
        {
            var normalizedFillColor = ParseHelpers.NormalizeArgbColor(cfNewFillColor);
            cfNewDxf.Append(new Fill(new PatternFill(
                new BackgroundColor { Rgb = normalizedFillColor })
            { PatternType = PatternValues.Solid }));
            cfNewHasDxf = true;
        }
        if (properties.TryGetValue("font.color", out _) && properties.TryGetValue("font.bold", out var cfNewFb2) && IsTruthy(cfNewFb2))
        {
            var existingFont = cfNewDxf.GetFirstChild<Font>();
            existingFont?.Append(new Bold());
        }

        if (cfNewHasDxf)
        {
            var cfNewWbPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var cfNewStyleMgr = new ExcelStyleManager(cfNewWbPart);
            cfNewStyleMgr.EnsureStylesPart();
            var cfNewStylesheet = cfNewWbPart.WorkbookStylesPart!.Stylesheet!;
            var cfNewDxfs = cfNewStylesheet.GetFirstChild<DifferentialFormats>();
            if (cfNewDxfs == null)
            {
                cfNewDxfs = new DifferentialFormats { Count = 0 };
                cfNewStylesheet.Append(cfNewDxfs);
            }
            cfNewDxfs.Append(cfNewDxf);
            cfNewDxfs.Count = (uint)cfNewDxfs.Elements<DifferentialFormat>().Count();
            _dirtyStylesheet = true;
            cfNewRule.FormatId = cfNewDxfs.Count!.Value - 1;
        }

        var cfNewFormatting = new ConditionalFormatting(cfNewRule)
        {
            SequenceOfReferences = new ListValue<StringValue>(
                cfNewSqref.Split(' ').Select(s => new StringValue(s)))
        };

        var cfNewWs = GetSheet(cfNewWorksheet);
        InsertConditionalFormatting(cfNewWs, cfNewFormatting);

        SaveWorksheet(cfNewWorksheet);
        var cfNewCount = cfNewWs.Elements<ConditionalFormatting>().Count();
        return $"/{cfNewSheetName}/cf[{cfNewCount}]";
    }

}
