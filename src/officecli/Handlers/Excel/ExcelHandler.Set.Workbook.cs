// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Try to handle workbook-level settings. Returns true if handled.
    /// </summary>
    private bool TrySetWorkbookSetting(string key, string value)
    {
        switch (key)
        {
            // ==================== WorkbookProperties ====================
            case "workbook.date1904" or "date1904":
            {
                var props = EnsureWorkbookProperties();
                props.Date1904 = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "workbook.codename" or "codename":
            {
                var props = EnsureWorkbookProperties();
                props.CodeName = value;
                SaveWorkbook();
                return true;
            }
            case "workbook.filterprivacy" or "filterprivacy":
            {
                var props = EnsureWorkbookProperties();
                props.FilterPrivacy = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "workbook.showobjects" or "showobjects":
            {
                var props = EnsureWorkbookProperties();
                props.ShowObjects = value.ToLowerInvariant() switch
                {
                    "all" => ObjectDisplayValues.All,
                    "placeholders" => ObjectDisplayValues.Placeholders,
                    "none" => ObjectDisplayValues.None,
                    _ => throw new ArgumentException($"Invalid showObjects: '{value}'. Valid: all, placeholders, none")
                };
                SaveWorkbook();
                return true;
            }
            case "workbook.backupfile" or "backupfile":
            {
                var props = EnsureWorkbookProperties();
                props.BackupFile = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "workbook.datecompatibility" or "datecompatibility":
            {
                var props = EnsureWorkbookProperties();
                props.DateCompatibility = IsTruthy(value);
                SaveWorkbook();
                return true;
            }

            // ==================== CalculationProperties ====================
            case "calc.mode" or "calcmode":
            {
                var calc = EnsureCalculationProperties();
                calc.CalculationMode = value.ToLowerInvariant() switch
                {
                    "auto" or "automatic" => CalculateModeValues.Auto,
                    "manual" => CalculateModeValues.Manual,
                    "autonoexcepttables" or "autoexcepttables" or "autonotable" => CalculateModeValues.AutoNoTable,
                    _ => throw new ArgumentException($"Invalid calc.mode: '{value}'. Valid: auto, manual, autoExceptTables")
                };
                SaveWorkbook();
                return true;
            }
            case "calc.iterate" or "iterate":
            {
                var calc = EnsureCalculationProperties();
                calc.Iterate = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "calc.iteratecount" or "iteratecount":
            {
                var calc = EnsureCalculationProperties();
                calc.IterateCount = ParseHelpers.SafeParseUint(value, "calc.iterateCount");
                SaveWorkbook();
                return true;
            }
            case "calc.iteratedelta" or "iteratedelta":
            {
                var calc = EnsureCalculationProperties();
                calc.IterateDelta = ParseHelpers.SafeParseDouble(value, "calc.iterateDelta");
                SaveWorkbook();
                return true;
            }
            case "calc.fullprecision" or "fullprecision":
            {
                var calc = EnsureCalculationProperties();
                calc.FullPrecision = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "calc.fullcalconload" or "fullcalconload":
            {
                var calc = EnsureCalculationProperties();
                calc.FullCalculationOnLoad = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "calc.refmode" or "refmode":
            {
                var calc = EnsureCalculationProperties();
                calc.ReferenceMode = value.ToLowerInvariant() switch
                {
                    "a1" => ReferenceModeValues.A1,
                    "r1c1" => ReferenceModeValues.R1C1,
                    _ => throw new ArgumentException($"Invalid calc.refMode: '{value}'. Valid: A1, R1C1")
                };
                SaveWorkbook();
                return true;
            }

            // ==================== WorkbookProtection ====================
            case "workbook.protection" or "workbookprotection":
            {
                var workbook = _doc.WorkbookPart!.Workbook;
                var existing = workbook.GetFirstChild<WorkbookProtection>();
                existing?.Remove();
                if (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase) && IsTruthy(value))
                {
                    var newProt = new WorkbookProtection { LockStructure = true, LockWindows = true };
                    var anchor = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<BookViews>()
                        ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<Sheets>()
                        ?? workbook.GetFirstChild<CalculationProperties>();
                    if (anchor != null)
                        anchor.InsertBeforeSelf(newProt);
                    else
                        workbook.AppendChild(newProt);
                }
                SaveWorkbook();
                return true;
            }
            case "workbook.lockstructure" or "lockstructure":
            {
                var prot = EnsureWorkbookProtection();
                prot.LockStructure = IsTruthy(value);
                SaveWorkbook();
                return true;
            }
            case "workbook.lockwindows" or "lockwindows":
            {
                var prot = EnsureWorkbookProtection();
                prot.LockWindows = IsTruthy(value);
                SaveWorkbook();
                return true;
            }

            default:
                return false;
        }
    }

    // ==================== Helpers ====================

    private WorkbookProperties EnsureWorkbookProperties()
    {
        var workbook = _doc.WorkbookPart!.Workbook;
        var props = workbook.GetFirstChild<WorkbookProperties>();
        if (props == null)
        {
            props = new WorkbookProperties();
            // Schema order: workbookPr must appear before Sheets, BookViews, etc.
            // Insert as the first child to maintain schema order.
            var firstChild = workbook.FirstChild;
            if (firstChild != null)
                firstChild.InsertBeforeSelf(props);
            else
                workbook.AppendChild(props);
        }
        return props;
    }

    private CalculationProperties EnsureCalculationProperties()
    {
        var workbook = _doc.WorkbookPart!.Workbook;
        var calc = workbook.GetFirstChild<CalculationProperties>();
        if (calc == null)
        {
            calc = new CalculationProperties();
            workbook.AppendChild(calc);
        }
        return calc;
    }

    private WorkbookProtection EnsureWorkbookProtection()
    {
        var workbook = _doc.WorkbookPart!.Workbook;
        var prot = workbook.GetFirstChild<WorkbookProtection>();
        if (prot == null)
        {
            prot = new WorkbookProtection();
            // Schema order: workbookProtection must precede bookViews and sheets.
            // Insert before the first of BookViews, Sheets, or CalculationProperties if present.
            var anchor = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<BookViews>()
                ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<Sheets>()
                ?? workbook.GetFirstChild<CalculationProperties>();
            if (anchor != null)
                anchor.InsertBeforeSelf(prot);
            else
                workbook.AppendChild(prot);
        }
        return prot;
    }

    private void SaveWorkbook()
    {
        _doc.WorkbookPart?.Workbook?.Save();
    }

    /// <summary>
    /// Read workbook-level settings into Format dictionary.
    /// </summary>
    private void PopulateWorkbookSettings(DocumentNode node)
    {
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return;

        // WorkbookProperties
        var props = workbook.GetFirstChild<WorkbookProperties>();
        if (props != null)
        {
            if (props.Date1904?.Value == true) node.Format["workbook.date1904"] = true;
            if (props.CodeName?.Value != null) node.Format["workbook.codeName"] = props.CodeName.Value;
            if (props.FilterPrivacy?.Value == true) node.Format["workbook.filterPrivacy"] = true;
            if (props.ShowObjects?.Value != null) node.Format["workbook.showObjects"] = props.ShowObjects.InnerText;
            if (props.BackupFile?.Value == true) node.Format["workbook.backupFile"] = true;
            if (props.DateCompatibility?.Value == true) node.Format["workbook.dateCompatibility"] = true;
        }

        // CalculationProperties
        var calc = workbook.GetFirstChild<CalculationProperties>();
        if (calc != null)
        {
            if (calc.CalculationMode?.Value != null) node.Format["calc.mode"] = calc.CalculationMode.InnerText;
            if (calc.Iterate?.Value == true) node.Format["calc.iterate"] = true;
            if (calc.IterateCount?.Value != null) node.Format["calc.iterateCount"] = (int)calc.IterateCount.Value;
            if (calc.IterateDelta?.Value != null) node.Format["calc.iterateDelta"] = calc.IterateDelta.Value;
            if (calc.FullPrecision?.Value != null) node.Format["calc.fullPrecision"] = calc.FullPrecision.Value;
            if (calc.FullCalculationOnLoad?.Value == true) node.Format["calc.fullCalcOnLoad"] = true;
            if (calc.ReferenceMode?.Value != null) node.Format["calc.refMode"] = calc.ReferenceMode.InnerText;
        }

        // WorkbookProtection
        var prot = workbook.GetFirstChild<WorkbookProtection>();
        if (prot != null)
        {
            if (prot.LockStructure?.Value == true) node.Format["workbook.lockStructure"] = true;
            if (prot.LockWindows?.Value == true) node.Format["workbook.lockWindows"] = true;
        }
    }
}
