// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Form Fields ====================

    /// <summary>
    /// Find all legacy form fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN) in the document.
    /// </summary>
    private List<(FieldInfo Field, FormFieldData FfData)> FindFormFields()
    {
        var allFields = FindFields();
        var result = new List<(FieldInfo, FormFieldData)>();
        foreach (var field in allFields)
        {
            var beginChar = field.BeginRun.GetFirstChild<FieldChar>();
            var ffData = beginChar?.FormFieldData;
            if (ffData != null)
                result.Add((field, ffData));
        }
        return result;
    }

    /// <summary>
    /// Convert a form field to a DocumentNode.
    /// </summary>
    private DocumentNode FormFieldToNode((FieldInfo Field, FormFieldData FfData) ff, string path)
    {
        var node = new DocumentNode { Path = path, Type = "formfield" };
        var ffData = ff.FfData;

        // Name
        var name = ffData.GetFirstChild<FormFieldName>()?.Val?.Value;
        if (name != null) node.Format["name"] = name;

        // Enabled
        var enabled = ffData.GetFirstChild<Enabled>();
        node.Format["enabled"] = enabled?.Val?.Value ?? true;

        // Determine formfield type and read type-specific properties
        var textInput = ffData.GetFirstChild<TextInput>();
        var checkBox = ffData.GetFirstChild<CheckBox>();
        var dropDown = ffData.GetFirstChild<DropDownListFormField>();

        if (textInput != null)
        {
            node.Format["formfieldType"] = "text";
            var defaultVal = textInput.GetFirstChild<DefaultTextBoxFormFieldString>()?.Val?.Value;
            if (defaultVal != null) node.Format["default"] = defaultVal;
            var maxLen = textInput.GetFirstChild<MaxLength>()?.Val?.Value;
            if (maxLen != null) node.Format["maxLength"] = (int)maxLen;
            // Result text (current value)
            var resultText = string.Join("", ff.Field.ResultRuns.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
            node.Text = resultText;
        }
        else if (checkBox != null)
        {
            node.Format["formfieldType"] = "checkbox";
            var checkedEl = checkBox.GetFirstChild<Checked>();
            var defaultEl = checkBox.GetFirstChild<DefaultCheckBoxFormFieldState>();
            var isChecked = checkedEl?.Val?.Value ?? defaultEl?.Val?.Value ?? false;
            node.Format["checked"] = isChecked;
            node.Text = isChecked ? "true" : "false";
        }
        else if (dropDown != null)
        {
            node.Format["formfieldType"] = "dropdown";
            var items = dropDown.Elements<ListEntryFormField>().Select(li => li.Val?.Value ?? "").ToList();
            if (items.Count > 0) node.Format["items"] = string.Join(",", items);
            var defaultIdx = dropDown.GetFirstChild<DropDownListSelection>()?.Val?.Value ?? 0;
            node.Format["default"] = (int)defaultIdx;
            // Current selection
            var resultText = string.Join("", ff.Field.ResultRuns.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
            node.Text = resultText;
            if (string.IsNullOrEmpty(resultText) && defaultIdx < items.Count)
                node.Text = items[(int)defaultIdx];
        }

        // Editable status based on protection
        node.Format["editable"] = IsFormFieldEditable(ffData);

        return node;
    }

    /// <summary>
    /// Check if a form field is editable based on document protection.
    /// </summary>
    private bool IsFormFieldEditable(FormFieldData ffData)
    {
        var (mode, enforced) = GetDocumentProtection();

        // No protection → editable
        if (!enforced || mode == "none")
            return true;

        // Forms protection → form fields are always editable (unless disabled)
        if (mode == "forms")
        {
            var enabled = ffData.GetFirstChild<Enabled>();
            return enabled?.Val?.Value ?? true;
        }

        // readOnly → not editable
        return false;
    }

    /// <summary>
    /// Set properties on a form field.
    /// </summary>
    private List<string> SetFormField((FieldInfo Field, FormFieldData FfData) ff, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ffData = ff.FfData;

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text" or "value":
                {
                    var textInput = ffData.GetFirstChild<TextInput>();
                    var checkBox = ffData.GetFirstChild<CheckBox>();
                    var dropDown = ffData.GetFirstChild<DropDownListFormField>();

                    if (checkBox != null)
                    {
                        // Set checkbox state
                        var isChecked = ParseHelpers.IsTruthy(value);
                        var checkedEl = checkBox.GetFirstChild<Checked>();
                        if (checkedEl != null) checkedEl.Val = new OnOffValue(isChecked);
                        else checkBox.AppendChild(new Checked { Val = new OnOffValue(isChecked) });

                        // Update result text (Word uses special checkbox symbol)
                        SetFormFieldResultText(ff.Field, isChecked ? "\u2612" : "\u2610");
                    }
                    else if (dropDown != null)
                    {
                        // Set dropdown selection by text or index
                        var items = dropDown.Elements<ListEntryFormField>().Select(li => li.Val?.Value ?? "").ToList();
                        int idx;
                        if (int.TryParse(value, out idx))
                        {
                            // By index
                            if (idx >= 0 && idx < items.Count)
                            {
                                var selEl = dropDown.GetFirstChild<DropDownListSelection>();
                                if (selEl != null) selEl.Val = idx;
                                else dropDown.AppendChild(new DropDownListSelection { Val = idx });
                                SetFormFieldResultText(ff.Field, items[idx]);
                            }
                        }
                        else
                        {
                            // By text match
                            var matchIdx = items.FindIndex(i => string.Equals(i, value, StringComparison.OrdinalIgnoreCase));
                            if (matchIdx >= 0)
                            {
                                var selEl = dropDown.GetFirstChild<DropDownListSelection>();
                                if (selEl != null) selEl.Val = matchIdx;
                                else dropDown.AppendChild(new DropDownListSelection { Val = matchIdx });
                                SetFormFieldResultText(ff.Field, items[matchIdx]);
                            }
                            else
                            {
                                SetFormFieldResultText(ff.Field, value);
                            }
                        }
                    }
                    else
                    {
                        // Text input - just replace result text
                        SetFormFieldResultText(ff.Field, value);
                    }
                    break;
                }
                case "checked":
                {
                    var checkBox = ffData.GetFirstChild<CheckBox>();
                    if (checkBox != null)
                    {
                        var isChecked = ParseHelpers.IsTruthy(value);
                        var checkedEl = checkBox.GetFirstChild<Checked>();
                        if (checkedEl != null) checkedEl.Val = new OnOffValue(isChecked);
                        else checkBox.AppendChild(new Checked { Val = new OnOffValue(isChecked) });
                        SetFormFieldResultText(ff.Field, isChecked ? "\u2612" : "\u2610");
                    }
                    else
                        unsupported.Add(key);
                    break;
                }
                case "name":
                {
                    var nameEl = ffData.GetFirstChild<FormFieldName>();
                    if (nameEl != null) nameEl.Val = value;
                    else ffData.PrependChild(new FormFieldName { Val = value });
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    /// <summary>
    /// Replace the result text of a form field (runs between separate and end).
    /// </summary>
    private static void SetFormFieldResultText(FieldInfo field, string text)
    {
        if (field.SeparateRun == null) return;

        // Remove existing result runs
        foreach (var run in field.ResultRuns)
            run.Remove();
        field.ResultRuns.Clear();

        // Insert new result run after the separate fieldchar run
        var newRun = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        // Copy run properties from the separate run or begin run for consistent formatting
        var sourceProps = field.SeparateRun.RunProperties ?? field.BeginRun.RunProperties;
        if (sourceProps != null)
            newRun.PrependChild(sourceProps.CloneNode(true));

        field.SeparateRun.InsertAfterSelf(newRun);
    }

    /// <summary>
    /// Add a legacy form field to a paragraph.
    /// </summary>
    private string AddFormField(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        Paragraph para;
        if (parent is Paragraph p)
        {
            para = p;
        }
        else if (parent is Body bodyEl)
        {
            para = new Paragraph();
            // Honor index (ChildElements-based) and the Body's trailing sectPr
            // — raw AppendChild put the paragraph AFTER sectPr, making the
            // document schema-invalid.
            InsertAtIndexOrAppend(bodyEl, para, index);
            // index was consumed by the placement above; clear it so the
            // later FormField re-threading (which also inspects index)
            // doesn't try to rearrange runs inside the new paragraph.
            index = null;
            var paraIdx = bodyEl.Elements<Paragraph>().ToList().IndexOf(para) + 1;
            parentPath = $"/body/{BuildParaPathSegment(para, paraIdx)}";
        }
        else
        {
            throw new ArgumentException("Form fields must be added to a paragraph or /body");
        }

        var ciProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        var ffType = ciProps.GetValueOrDefault("formfieldtype",
            ciProps.GetValueOrDefault("type", "text")).ToLowerInvariant();
        // Treat explicit name="" the same as missing name: auto-generate.
        // Empty bookmark names are addressable-invalid (predicate validator
        // rejects bare empty values), and the validator below would crash
        // on name[0] if we let "" through.
        var name = ciProps.GetValueOrDefault("name", "");
        if (string.IsNullOrEmpty(name))
            name = $"ff_{Guid.NewGuid():N}"[..12];
        if (name.Any(c => c == '/' || c == '[' || c == ']'))
            throw new ArgumentException(
                $"Form field name '{name}' contains path-special characters " +
                "('/', '[', ']'). These characters prevent later addressing via " +
                "selectors. Use only letters, digits, '.', '_', '-' in form field names.");
        // Form fields embed a BookmarkStart/End with the same name, so they
        // must obey the same addressability rules as bookmarks (R18): no
        // whitespace, no leading '@'/'\'', no embedded '"', and no duplicate
        // names anywhere in the document.
        if (name.Any(char.IsWhiteSpace) || name[0] == '@' || name[0] == '\'' || name.Contains('"'))
            throw new ArgumentException(
                $"Form field name '{name}' contains whitespace or quote/@ chars " +
                "that prevent later addressing via bare attribute selectors. " +
                "Use only letters, digits, '.', '_', '-' in form field names.");
        if (body.Descendants<BookmarkStart>()
                .Any(b => string.Equals(b.Name?.Value, name, StringComparison.Ordinal)))
        {
            throw new ArgumentException(
                $"form field name '{name}' already exists as a bookmark; pick a unique name.");
        }
        var text = ciProps.GetValueOrDefault("text", ciProps.GetValueOrDefault("value", ""));

        // Generate unique bookmark ID
        var existingIds = body.Descendants<BookmarkStart>()
            .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
        var bkId = (existingIds.Any() ? existingIds.Max() + 1 : 1).ToString();

        // BookmarkStart
        var bookmarkStart = new BookmarkStart { Id = bkId, Name = name };
        para.AppendChild(bookmarkStart);

        // Begin run with FieldChar(Begin) + FormFieldData
        var beginRun = new Run();
        var beginChar = new FieldChar { FieldCharType = FieldCharValues.Begin };

        var ffData = new FormFieldData();
        ffData.AppendChild(new FormFieldName { Val = name });
        ffData.AppendChild(new Enabled());

        switch (ffType)
        {
            case "checkbox" or "check":
            {
                var checkBox = new CheckBox();
                checkBox.AppendChild(new FormFieldSize { Val = "20" }); // Default size in half-points
                var isChecked = ciProps.TryGetValue("checked", out var chkVal) && ParseHelpers.IsTruthy(chkVal);
                checkBox.AppendChild(new DefaultCheckBoxFormFieldState { Val = new OnOffValue(isChecked) });
                if (isChecked)
                    checkBox.AppendChild(new Checked { Val = new OnOffValue(true) });
                ffData.AppendChild(checkBox);
                text = isChecked ? "\u2612" : "\u2610";
                break;
            }
            case "dropdown" or "drop":
            {
                var ddl = new DropDownListFormField();
                if (ciProps.TryGetValue("items", out var items))
                {
                    foreach (var item in items.Split(','))
                        ddl.AppendChild(new ListEntryFormField { Val = item.Trim() });
                }
                ffData.AppendChild(ddl);
                // Default to first item if no text specified
                if (string.IsNullOrEmpty(text) && ciProps.TryGetValue("items", out var itemsList))
                {
                    var firstItem = itemsList.Split(',').FirstOrDefault()?.Trim();
                    if (firstItem != null) text = firstItem;
                }
                break;
            }
            default: // "text"
            {
                var textInput = new TextInput();
                if (ciProps.TryGetValue("default", out var defaultVal))
                {
                    textInput.AppendChild(new DefaultTextBoxFormFieldString { Val = defaultVal });
                    // Use default value as initial text if no explicit text/value provided
                    if (string.IsNullOrEmpty(text))
                        text = defaultVal;
                }
                if (ciProps.TryGetValue("maxlength", out var maxLenStr) && int.TryParse(maxLenStr, out var maxLen))
                    textInput.AppendChild(new MaxLength { Val = (short)maxLen });
                ffData.AppendChild(textInput);
                break;
            }
        }

        beginChar.AppendChild(ffData);
        beginRun.AppendChild(beginChar);
        para.AppendChild(beginRun);

        // Instruction run
        var instrText = ffType switch
        {
            "checkbox" or "check" => " FORMCHECKBOX ",
            "dropdown" or "drop" => " FORMDROPDOWN ",
            _ => " FORMTEXT "
        };
        var instrRun = new Run(new FieldCode(instrText) { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(instrRun);

        // Separate run
        var separateRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
        para.AppendChild(separateRun);

        // Result run
        if (!string.IsNullOrEmpty(text))
        {
            var resultRun = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(resultRun);
        }
        else
        {
            // Add default placeholder for FORMTEXT
            var resultRun = new Run(new Text("\u00A0") { Space = SpaceProcessingModeValues.Preserve }); // non-breaking space
            para.AppendChild(resultRun);
        }

        // End run
        var endRun = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
        para.AppendChild(endRun);

        // BookmarkEnd
        var bookmarkEnd = new BookmarkEnd { Id = bkId };
        para.AppendChild(bookmarkEnd);

        // CONSISTENCY(add-index): honor --index / --after / --before (#76).
        // When an anchor/index was supplied, re-thread the 7 appended elements
        // into the requested child-element position. Simpler than restructuring
        // the construction path above.
        if (index.HasValue)
        {
            // Snapshot: the 7 elements we just appended, in order.
            var ffElements = para.ChildElements
                .Reverse().Take(7).Reverse().ToList();
            // The anchor position was computed against the children BEFORE we
            // appended the 7 elements. Subtract those 7 from the current count
            // to get the original anchor child.
            var origChildCount = para.ChildElements.Count - ffElements.Count;
            if (index.Value < origChildCount)
            {
                var anchor = para.ChildElements[index.Value];
                foreach (var el in ffElements) el.Remove();
                para.InsertBefore(ffElements[0], anchor);
                for (int ffI = 1; ffI < ffElements.Count; ffI++)
                    para.InsertAfter(ffElements[ffI], ffElements[ffI - 1]);
            }
            // else: index is at or past the end — current append position is correct.
        }

        _doc.MainDocumentPart?.Document?.Save();

        // Compute result path
        int ffIdx = 0;
        var allFf = FindFormFields();
        for (int i = 0; i < allFf.Count; i++)
        {
            if (allFf[i].Field.BeginRun == beginRun)
            { ffIdx = i + 1; break; }
        }
        return $"/formfield[{ffIdx}]";
    }
}
