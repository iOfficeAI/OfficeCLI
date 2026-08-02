// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // True only when the user explicitly set workbook.lockStructure (the
    // `workbook.lockstructure` case below). A bare workbook.password Set
    // auto-implies lockStructure=true for UI parity, but that implied lock is
    // NOT explicit, so clearing the password undoes it. An explicitly requested
    // lock survives a password clear. Default false: anything we did not see an
    // explicit request for is treated as auto-implied and removed on clear.
    private bool _workbookLockStructureExplicit;

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
                if (IsTruthy(value)) props.Date1904 = true;
                else props.Date1904 = null;
                CleanupEmptyWorkbookProperties();
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
                if (IsTruthy(value)) props.FilterPrivacy = true;
                else props.FilterPrivacy = null;
                CleanupEmptyWorkbookProperties();
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
                if (IsTruthy(value)) props.BackupFile = true;
                else props.BackupFile = null;
                CleanupEmptyWorkbookProperties();
                SaveWorkbook();
                return true;
            }
            case "workbook.datecompatibility" or "datecompatibility":
            {
                var props = EnsureWorkbookProperties();
                if (IsTruthy(value))
                    props.DateCompatibility = true;
                else
                    props.DateCompatibility = null;
                CleanupEmptyWorkbookProperties();
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
                if (IsTruthy(value))
                    calc.Iterate = true;
                else
                    calc.Iterate = null;
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
                // OOXML default is true; must write explicit false to override.
                calc.FullPrecision = IsTruthy(value) ? null : false;
                SaveWorkbook();
                return true;
            }
            case "calc.fullcalconload" or "fullcalconload":
            {
                var calc = EnsureCalculationProperties();
                if (IsTruthy(value))
                    calc.FullCalculationOnLoad = true;
                else
                    calc.FullCalculationOnLoad = null;
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

            // ==================== BookViews / WorkbookView ====================
            case "activetab" or "workbook.activetab":
            {
                var bv = EnsureFirstWorkbookView();
                // Accept 0-based numeric index or sheet name.
                uint idx;
                if (uint.TryParse(value, System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                {
                    idx = parsed;
                }
                else
                {
                    var sheets = _doc.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()
                        ?.Elements<Sheet>().ToList();
                    if (sheets == null || sheets.Count == 0)
                        throw new ArgumentException($"Invalid activeTab: no sheets in workbook");
                    var match = sheets.FindIndex(s =>
                        string.Equals(s.Name?.Value, value, StringComparison.OrdinalIgnoreCase));
                    if (match < 0)
                        throw new ArgumentException(
                            $"Invalid activeTab: '{value}' is not a 0-based index or sheet name. " +
                            $"Valid sheets: {string.Join(", ", sheets.Select(s => s.Name?.Value))}");
                    idx = (uint)match;
                }
                bv.ActiveTab = idx == 0 ? null : new UInt32Value(idx);
                SaveWorkbook();
                return true;
            }
            case "firstsheet" or "workbook.firstsheet":
            {
                var bv = EnsureFirstWorkbookView();
                uint idx;
                if (uint.TryParse(value, System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                {
                    idx = parsed;
                }
                else
                {
                    var sheets = _doc.WorkbookPart?.Workbook?.GetFirstChild<Sheets>()
                        ?.Elements<Sheet>().ToList();
                    if (sheets == null || sheets.Count == 0)
                        throw new ArgumentException($"Invalid firstSheet: no sheets in workbook");
                    var match = sheets.FindIndex(s =>
                        string.Equals(s.Name?.Value, value, StringComparison.OrdinalIgnoreCase));
                    if (match < 0)
                        throw new ArgumentException(
                            $"Invalid firstSheet: '{value}' is not a 0-based index or sheet name.");
                    idx = (uint)match;
                }
                bv.FirstSheet = idx == 0 ? null : new UInt32Value(idx);
                SaveWorkbook();
                return true;
            }

            // ==================== WorkbookProtection ====================
            case "workbook.protection" or "workbookprotection":
            {
                var workbook = _doc.WorkbookPart!.Workbook!;
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
                if (IsTruthy(value))
                {
                    prot.LockStructure = true;
                    // Explicit user request — must survive a later password clear.
                    _workbookLockStructureExplicit = true;
                }
                else
                {
                    prot.LockStructure = null;
                    _workbookLockStructureExplicit = false;
                }
                CleanupEmptyWorkbookProtection();
                SaveWorkbook();
                return true;
            }
            case "workbook.lockwindows" or "lockwindows":
            {
                var prot = EnsureWorkbookProtection();
                if (IsTruthy(value))
                    prot.LockWindows = true;
                else
                    prot.LockWindows = null;
                CleanupEmptyWorkbookProtection();
                SaveWorkbook();
                return true;
            }
            case "workbook.passwordhash":
            {
                // Verbatim legacy hash write (round-trip of the readback above),
                // parallel to the sheet-level `passwordhash` case. No re-hash:
                // the dump carries the already-hashed value. Implies a lock so
                // the protection element is meaningful (Excel parity).
                var protH = EnsureWorkbookProtection();
                if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    protH.WorkbookPassword = null;
                    if (!_workbookLockStructureExplicit) protH.LockStructure = null;
                }
                else
                {
                    protH.WorkbookPassword = HexBinaryValue.FromString(value);
                    if (protH.LockStructure?.Value != true && protH.LockWindows?.Value != true)
                        protH.LockStructure = true;
                }
                CleanupEmptyWorkbookProtection();
                SaveWorkbook();
                return true;
            }
            case "workbook.password" or "workbookpassword":
            {
                var prot = EnsureWorkbookProtection();
                if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    prot.WorkbookPassword = null;
                    // Undo the lockStructure that setting the password auto-implied, so a
                    // cleared-password workbook has no leftover lock. Only an explicitly
                    // requested lockStructure (the `workbook.lockstructure` case) survives.
                    // We key off "was it explicitly requested" rather than "did we imply
                    // it" so the clear is robust even when the implied lock came from a
                    // path that did not run the auto-imply branch (e.g. lockWindows was
                    // already present, or the element was loaded pre-locked).
                    if (!_workbookLockStructureExplicit)
                        prot.LockStructure = null;
                }
                else
                {
                    // ECMA-376 Part 4 14.7.1 legacy password hash (same algorithm
                    // used by sheet password). Truncated to 16-bit short — known
                    // weak, but matches what Excel writes for back-compat password
                    // fields without the modern algorithmName/saltValue/hashValue
                    // triple.
                    int hash = 0;
                    for (int ci = value.Length - 1; ci >= 0; ci--)
                    {
                        hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
                        hash ^= value[ci];
                    }
                    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
                    hash ^= value.Length;
                    hash ^= 0xCE4B;
                    prot.WorkbookPassword = HexBinaryValue.FromString(hash.ToString("X4"));
                    // Implies lockStructure unless caller overrides — mirrors Excel UI
                    // (the password field is only meaningful with at least one lock).
                    // This lock is auto-implied (not explicit), so a later password
                    // clear removes it — see the clear branch above.
                    if (prot.LockStructure?.Value != true && prot.LockWindows?.Value != true)
                        prot.LockStructure = true;
                }
                CleanupEmptyWorkbookProtection();
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
        var workbook = _doc.WorkbookPart!.Workbook!;
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
        var workbook = _doc.WorkbookPart!.Workbook!;
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
        var workbook = _doc.WorkbookPart!.Workbook!;
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

    private WorkbookView EnsureFirstWorkbookView()
    {
        var workbook = _doc.WorkbookPart!.Workbook!;
        var bookViews = workbook.GetFirstChild<BookViews>();
        if (bookViews == null)
        {
            bookViews = new BookViews();
            // Schema order: bookViews sits between workbookProtection/workbookPr
            // and sheets. Insert before Sheets when present.
            var anchor = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<Sheets>()
                ?? workbook.GetFirstChild<CalculationProperties>();
            if (anchor != null)
                anchor.InsertBeforeSelf(bookViews);
            else
                workbook.AppendChild(bookViews);
        }
        var view = bookViews.GetFirstChild<WorkbookView>();
        if (view == null)
        {
            view = new WorkbookView();
            bookViews.AppendChild(view);
        }
        return view;
    }

    private void CleanupEmptyWorkbookProperties()
    {
        var props = _doc.WorkbookPart?.Workbook?.GetFirstChild<WorkbookProperties>();
        if (props != null && !props.HasAttributes && !props.HasChildren)
            props.Remove();
    }

    private void CleanupEmptyWorkbookProtection()
    {
        var prot = _doc.WorkbookPart?.Workbook?.GetFirstChild<WorkbookProtection>();
        if (prot != null && !prot.HasAttributes && !prot.HasChildren)
            prot.Remove();
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

        // CalculationProperties — fullPrecision defaults to true per OOXML spec
        // even when the calc element is absent or attribute is omitted.
        var calc = workbook.GetFirstChild<CalculationProperties>();
        node.Format["calc.fullPrecision"] = calc?.FullPrecision?.Value ?? true;
        if (calc != null)
        {
            if (calc.CalculationMode?.Value != null) node.Format["calc.mode"] = calc.CalculationMode.InnerText;
            if (calc.Iterate?.Value == true) node.Format["calc.iterate"] = true;
            if (calc.IterateCount?.Value != null) node.Format["calc.iterateCount"] = (int)calc.IterateCount.Value;
            if (calc.IterateDelta?.Value != null) node.Format["calc.iterateDelta"] = calc.IterateDelta.Value;
            if (calc.FullCalculationOnLoad?.Value == true) node.Format["calc.fullCalcOnLoad"] = true;
            if (calc.ReferenceMode?.Value != null) node.Format["calc.refMode"] = calc.ReferenceMode.InnerText;
        }

        // BookViews / first WorkbookView
        var bookViews = workbook.GetFirstChild<BookViews>();
        var firstView = bookViews?.GetFirstChild<WorkbookView>();
        if (firstView != null)
        {
            if (firstView.ActiveTab?.Value is uint activeTab && activeTab != 0)
                node.Format["activeTab"] = (int)activeTab;
            if (firstView.FirstSheet?.Value is uint firstSheet && firstSheet != 0)
                node.Format["firstSheet"] = (int)firstSheet;
        }

        // WorkbookProtection
        var prot = workbook.GetFirstChild<WorkbookProtection>();
        if (prot != null)
        {
            if (prot.LockStructure?.Value == true) node.Format["workbook.lockStructure"] = true;
            if (prot.LockWindows?.Value == true) node.Format["workbook.lockWindows"] = true;
            // Surface the legacy password hash verbatim (canonical
            // `workbook.passwordHash`) so dump→batch preserves workbook
            // protection strength instead of degrading to passwordless —
            // mirrors the sheet-level passwordHash readback. `workbook.password`
            // stays a display-only mask (never round-tripped; the hash is).
            if (prot.WorkbookPassword?.Value is { Length: > 0 } wbPwHash)
            {
                node.Format["workbook.passwordHash"] = wbPwHash;
                node.Format["workbook.password"] = "***";
            }
        }
    }
}
