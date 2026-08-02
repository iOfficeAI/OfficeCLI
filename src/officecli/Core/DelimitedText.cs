// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Quote-aware grid parser for the table `data=` property, in both of its
/// forms: inline ("H1,H2;R1C1,R1C2" — semicolon rows) and CSV pulled from a
/// file/URL/data-URI (newline rows).
///
/// Splitting on the separator alone loses any cell that contains one, which is
/// exactly what quoting exists for: "Doe, John",30 is two fields, not three.
///
/// Excel import keeps its own parser (ExcelHandler.ParseCsv) — it has to
/// preserve blank source lines to keep row numbers aligned with the sheet,
/// while a table built from `data=` wants empty rows dropped.
/// </summary>
public static class DelimitedText
{
    /// <summary>
    /// Split <paramref name="content"/> into rows of fields. A field wrapped in
    /// double quotes may contain the field separator, the row separator and
    /// doubled quotes ("" → a literal "). Unquoted fields are trimmed; quoted
    /// fields keep their interior verbatim. Rows whose fields are all empty are
    /// dropped, so a trailing newline does not add a phantom row. When
    /// <paramref name="rowSeparator"/> is '\n', CRLF is handled too.
    /// </summary>
    public static string[][] ParseGrid(string content, char fieldSeparator, char rowSeparator)
    {
        var rows = new List<string[]>();
        if (string.IsNullOrEmpty(content)) return rows.ToArray();
        if (content[0] == '﻿') content = content[1..];

        var row = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        var quoted = false;

        void EndField()
        {
            row.Add(quoted ? field.ToString() : field.ToString().Trim());
            field.Clear();
            quoted = false;
        }

        void EndRow()
        {
            EndField();
            if (row.Exists(c => c.Length > 0)) rows.Add(row.ToArray());
            row.Clear();
        }

        for (int i = 0; i < content.Length; i++)
        {
            var c = content[i];

            if (inQuotes)
            {
                if (c != '"') { field.Append(c); continue; }
                // "" inside a quoted field is one literal quote; a lone quote closes it.
                if (i + 1 < content.Length && content[i + 1] == '"') { field.Append('"'); i++; }
                else inQuotes = false;
                continue;
            }

            if (c == '"' && IsBlank(field))
            {
                inQuotes = true;
                quoted = true;
                field.Clear();          // drop the whitespace that preceded the quote
                continue;
            }
            if (c == fieldSeparator) { EndField(); continue; }
            if (c == rowSeparator) { EndRow(); continue; }
            if (rowSeparator == '\n' && c == '\r') continue;
            field.Append(c);
        }

        EndRow();
        return rows.ToArray();
    }

    private static bool IsBlank(StringBuilder sb)
    {
        for (int i = 0; i < sb.Length; i++)
            if (!char.IsWhiteSpace(sb[i])) return false;
        return true;
    }
}
