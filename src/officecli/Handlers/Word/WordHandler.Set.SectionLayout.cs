// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Set section-level layout properties: Columns, SectionType.
    /// Called from TrySetDocSetting for keys with recognized prefixes.
    /// Returns true if the key was handled.
    /// </summary>
    private bool TrySetSectionLayout(string key, string value)
    {
        switch (key)
        {
            // ==================== Columns ====================
            case "columns.count":
            {
                var cols = EnsureColumns();
                cols.ColumnCount = (short)ParseHelpers.SafeParseInt(value, "columns.count");
                if (cols.EqualWidth == null)
                    cols.EqualWidth = true;
                return true;
            }
            // CONSISTENCY(canonical-key): 'columnSpace' is the canonical key
            // returned by Get/Query (see WordHandler.Query.cs:491); accept it
            // alongside the dotted alias so Set has parity with the read side.
            case "columns.space" or "columnspace":
            {
                var cols = EnsureColumns();
                cols.Space = ParseTwips(value).ToString();
                return true;
            }
            case "columns.equalwidth":
            {
                var cols = EnsureColumns();
                cols.EqualWidth = IsTruthy(value);
                return true;
            }
            case "columns.separator":
            {
                var cols = EnsureColumns();
                cols.Separator = IsTruthy(value);
                return true;
            }

            // ==================== SectionType ====================
            case "section.type" or "sectiontype":
            {
                var sectPr = EnsureSectionProperties();
                var sectType = sectPr.GetFirstChild<SectionType>();
                if (sectType == null)
                {
                    sectType = new SectionType();
                    sectPr.PrependChild(sectType);
                }
                sectType.Val = value.ToLowerInvariant() switch
                {
                    "nextpage" or "next" => SectionMarkValues.NextPage,
                    "continuous" => SectionMarkValues.Continuous,
                    "evenpage" or "even" => SectionMarkValues.EvenPage,
                    "oddpage" or "odd" => SectionMarkValues.OddPage,
                    "nextcolumn" or "column" => SectionMarkValues.NextColumn,
                    _ => throw new ArgumentException($"Invalid section.type: '{value}'. Valid: nextPage, continuous, evenPage, oddPage, nextColumn")
                };
                return true;
            }

            default:
                return false;
        }
    }

    private Columns EnsureColumns()
    {
        var sectPr = EnsureSectionProperties();
        var cols = sectPr.GetFirstChild<Columns>();
        if (cols == null)
        {
            cols = new Columns();
            // Schema order: cols must come before docGrid
            var docGrid = sectPr.GetFirstChild<DocGrid>();
            if (docGrid != null)
                docGrid.InsertBeforeSelf(cols);
            else
                sectPr.AppendChild(cols);
        }
        return cols;
    }
}
