// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Parses CSS-like attribute filters from query selectors and matches them against DocumentNode.
/// Supports operators: = (exact), != (not equal), ~= (contains), >= (greater or equal), <= (less or equal).
/// Example: "shape[fill=#FF0000][size>=24pt][text~=报告]"
/// </summary>
public static class AttributeFilter
{
    public enum FilterOp { Equal, NotEqual, Contains, GreaterOrEqual, LessOrEqual, GreaterThan, LessThan, Exists }

    public record Condition(string Key, FilterOp Op, string Value);

    // Regex: [key op value] where op is ~=, >=, <=, !=, =, >, or <
    // Order matters: multi-char operators before single-char to avoid partial match
    private static readonly Regex AttrRegex = new(
        @"\[([\w.]+)\s*(~=|>=|<=|\\?!=|=|>|<)\s*([^\]]*)\]",
        RegexOptions.Compiled);

    // Regex: [key] (has-attribute, no operator)
    private static readonly Regex HasAttrRegex = new(
        @"\[([\w.]+)\]",
        RegexOptions.Compiled);

    // Regex to find any [...] block (for validation)
    private static readonly Regex BracketBlockRegex = new(
        @"\[([^\]]*)\]",
        RegexOptions.Compiled);

    /// <summary>
    /// Parse all [key op value] conditions from a selector string.
    /// Throws CliException for malformed selectors.
    /// </summary>
    public static List<Condition> Parse(string selector)
    {
        // Check for unclosed brackets
        var openCount = selector.Count(c => c == '[');
        var closeCount = selector.Count(c => c == ']');
        if (openCount != closeCount)
            throw new CliException($"Malformed selector: unclosed bracket in \"{selector}\"")
            {
                Code = "invalid_selector",
                Suggestion = "Ensure every '[' has a matching ']'. Example: paragraph[style=Heading 1]"
            };

        var conditions = new List<Condition>();
        var matchedSpans = new HashSet<(int Start, int End)>();

        foreach (Match m in AttrRegex.Matches(selector))
        {
            var key = m.Groups[1].Value;
            var opStr = m.Groups[2].Value.Replace("\\", "");
            var val = m.Groups[3].Value.Trim('\'', '"');

            // Detect corrupted values from mis-parsed operators (e.g. === parsed as = with value ==X)
            if (val.StartsWith("=") || val.StartsWith("~") || val.StartsWith("!"))
                throw new CliException($"Malformed selector: invalid operator in \"[{m.Groups[0].Value.Trim('[', ']')}]\". Supported operators: =, !=, ~=, >=, <=, >, <")
                {
                    Code = "invalid_selector",
                    Suggestion = $"Did you mean [{key}={val.TrimStart('=', '~', '!')}]?"
                };

            var op = opStr switch
            {
                "~=" => FilterOp.Contains,
                ">=" => FilterOp.GreaterOrEqual,
                "<=" => FilterOp.LessOrEqual,
                ">" => FilterOp.GreaterThan,
                "<" => FilterOp.LessThan,
                "!=" => FilterOp.NotEqual,
                _ => FilterOp.Equal
            };

            conditions.Add(new Condition(key, op, val));
            matchedSpans.Add((m.Index, m.Index + m.Length));
        }

        // Find [...] blocks that weren't matched by the key=value regex
        foreach (Match block in BracketBlockRegex.Matches(selector))
        {
            if (matchedSpans.Any(s => s.Start == block.Index)) continue;
            var content = block.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(content))
                throw new CliException($"Malformed selector: empty brackets \"[]\" in \"{selector}\"")
                {
                    Code = "invalid_selector",
                    Suggestion = "Use [key=value] or [key] syntax. Example: paragraph[style=Heading 1]"
                };
            // Index like [1] — valid path syntax, skip
            if (int.TryParse(content, out _)) continue;
            // [key] with no operator — "has attribute" filter (CSS [attr] syntax)
            var hasAttrMatch = HasAttrRegex.Match(block.Value);
            if (hasAttrMatch.Success)
            {
                conditions.Add(new Condition(hasAttrMatch.Groups[1].Value, FilterOp.Exists, ""));
                matchedSpans.Add((block.Index, block.Index + block.Length));
                continue;
            }
            // Unrecognized bracket content
            throw new CliException($"Malformed selector: cannot parse \"[{content}]\". Expected [key=value] with operator =, !=, ~=, >=, <=, >, or <")
            {
                Code = "invalid_selector",
                Suggestion = "Example: paragraph[style=Heading 1], shape[fill!=#FF0000], cell[formula]"
            };
        }

        return conditions;
    }

    /// <summary>
    /// Filter a list of DocumentNodes by the given conditions.
    /// All operators (=, !=, ~=, >=, <=) are applied as a post-filter.
    /// This is safe even when handler selectors already pre-filter = and !=,
    /// since filtering is idempotent.
    /// </summary>
    public static List<DocumentNode> Apply(List<DocumentNode> nodes, List<Condition> conditions, bool applyAll = true)
    {
        if (conditions.Count == 0) return nodes;

        var toApply = applyAll
            ? conditions
            : conditions.Where(c => c.Op is FilterOp.Contains or FilterOp.GreaterOrEqual or FilterOp.LessOrEqual or FilterOp.GreaterThan or FilterOp.LessThan or FilterOp.Exists).ToList();

        if (toApply.Count == 0) return nodes;

        return nodes.Where(n => MatchAll(n, toApply)).ToList();
    }

    /// <summary>
    /// Filter nodes and collect diagnostic warnings.
    /// Warns when: a filter key doesn't exist in ANY node's Format,
    /// or when >= / <= / > / < is used on a non-numeric value.
    /// </summary>
    public static (List<DocumentNode> Results, List<string> Warnings) ApplyWithWarnings(
        List<DocumentNode> nodes, List<Condition> conditions, bool applyAll = true)
    {
        var warnings = new List<string>();
        if (conditions.Count == 0) return (nodes, warnings);

        var toApply = applyAll
            ? conditions
            : conditions.Where(c => c.Op is FilterOp.Contains or FilterOp.GreaterOrEqual or FilterOp.LessOrEqual or FilterOp.GreaterThan or FilterOp.LessThan or FilterOp.Exists).ToList();

        if (toApply.Count == 0) return (nodes, warnings);

        // Check for missing keys: if a filter key doesn't exist in ANY node, warn
        foreach (var cond in toApply)
        {
            if (cond.Op == FilterOp.NotEqual) continue; // missing key is valid for !=
            bool anyHasKey = nodes.Any(n => ResolveValue(n, cond.Key).HasKey);
            if (!anyHasKey && nodes.Count > 0)
            {
                warnings.Add($"Warning: filter key '{cond.Key}' not found in any result's Format. " +
                    $"Available keys: {string.Join(", ", GetAllFormatKeys(nodes))}");
            }
        }

        // Check for non-numeric values on >= / <= / > / <
        foreach (var cond in toApply.Where(c => c.Op is FilterOp.GreaterOrEqual or FilterOp.LessOrEqual or FilterOp.GreaterThan or FilterOp.LessThan))
        {
            if (ExtractNumber(cond.Value) == null && !EmuConverter.TryParseEmu(cond.Value, out _))
            {
                warnings.Add($"Warning: '{cond.Value}' in [{cond.Key}{OpToString(cond.Op)}{cond.Value}] " +
                    $"is not numeric — comparison may produce unexpected results");
            }
            // Also check actual values in nodes
            foreach (var node in nodes)
            {
                var (hasKey, actual) = ResolveValue(node, cond.Key);
                if (hasKey && ExtractNumber(actual) == null && !EmuConverter.TryParseEmu(actual, out _))
                {
                    warnings.Add($"Warning: value '{actual}' for key '{cond.Key}' at {node.Path} " +
                        $"is not numeric — {OpToString(cond.Op)} comparison may be unreliable");
                    break; // one warning per condition is enough
                }
            }
        }

        var results = nodes.Where(n => MatchAll(n, toApply)).ToList();
        return (results, warnings);
    }

    private static string OpToString(FilterOp op) => op switch
    {
        FilterOp.Equal => "=",
        FilterOp.NotEqual => "!=",
        FilterOp.Contains => "~=",
        FilterOp.GreaterOrEqual => ">=",
        FilterOp.LessOrEqual => "<=",
        FilterOp.GreaterThan => ">",
        FilterOp.LessThan => "<",
        FilterOp.Exists => "(exists)",
        _ => "?"
    };

    private static HashSet<string> GetAllFormatKeys(List<DocumentNode> nodes)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var node in nodes)
        {
            foreach (var key in node.Format.Keys)
                keys.Add(key);
            if (node.Text != null) keys.Add("text");
            if (!string.IsNullOrEmpty(node.Type)) keys.Add("type");
        }
        return keys;
    }

    /// <summary>
    /// Check if a DocumentNode matches all conditions.
    /// </summary>
    public static bool MatchAll(DocumentNode node, List<Condition> conditions)
    {
        foreach (var cond in conditions)
        {
            if (!MatchOne(node, cond)) return false;
        }
        return true;
    }

    private static bool MatchOne(DocumentNode node, Condition cond)
    {
        // Resolve actual value from node
        var (hasKey, actualStr) = ResolveValue(node, cond.Key);

        switch (cond.Op)
        {
            case FilterOp.Exists:
                return hasKey && !string.IsNullOrEmpty(actualStr);

            case FilterOp.Equal:
                if (!hasKey) return false;
                return StringEquals(actualStr, cond.Value)
                    || DimensionEquals(actualStr, cond.Value);

            case FilterOp.NotEqual:
                if (!hasKey) return true; // key absent → not equal
                return !StringEquals(actualStr, cond.Value)
                    && !DimensionEquals(actualStr, cond.Value);

            case FilterOp.Contains:
                if (!hasKey) return false;
                return actualStr.Contains(cond.Value, StringComparison.OrdinalIgnoreCase);

            case FilterOp.GreaterOrEqual:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) >= 0;

            case FilterOp.LessOrEqual:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) <= 0;

            case FilterOp.GreaterThan:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) > 0;

            case FilterOp.LessThan:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) < 0;

            default:
                return true;
        }
    }

    private static (bool HasKey, string Value) ResolveValue(DocumentNode node, string key)
    {
        // Case-insensitive Format key lookup (highest priority)
        var matchedKey = node.Format.Keys.FirstOrDefault(k =>
            string.Equals(k, key, StringComparison.OrdinalIgnoreCase));

        if (matchedKey != null)
        {
            var val = node.Format[matchedKey];
            return (true, val?.ToString() ?? "");
        }

        // "text" falls back to node.Text if not in Format
        if (string.Equals(key, "text", StringComparison.OrdinalIgnoreCase))
        {
            return (node.Text != null, node.Text ?? "");
        }

        // "type" falls back to node.Type if not in Format
        if (string.Equals(key, "type", StringComparison.OrdinalIgnoreCase))
        {
            return (!string.IsNullOrEmpty(node.Type), node.Type ?? "");
        }

        return (false, "");
    }

    private static bool StringEquals(string a, string b)
    {
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
            return true;
        // Normalize color hex: "#FF0000" matches "FF0000" and vice versa
        var aNorm = a.TrimStart('#');
        var bNorm = b.TrimStart('#');
        if (aNorm != a || bNorm != b)
            return string.Equals(aNorm, bNorm, StringComparison.OrdinalIgnoreCase);
        return false;
    }

    private static bool DimensionEquals(string actual, string expected)
    {
        if (EmuConverter.TryParseEmu(actual, out var a) && EmuConverter.TryParseEmu(expected, out var b))
            return Math.Abs(a - b) <= 500;
        return false;
    }

    /// <summary>
    /// Compare two values numerically. Supports:
    /// - Plain numbers: "24", "1.5"
    /// - pt-suffixed: "24pt", "10.5pt"
    /// - EMU/dimension values: "2cm", "1in"
    /// Returns negative if actual &lt; expected, 0 if equal, positive if actual &gt; expected.
    /// Falls back to string comparison if neither numeric nor dimension.
    /// </summary>
    private static int CompareNumeric(string actual, string expected)
    {
        // Try plain decimal comparison (handles "24", "1.5", "24pt" vs "20pt", etc.)
        var actualNum = ExtractNumber(actual);
        var expectedNum = ExtractNumber(expected);

        if (actualNum.HasValue && expectedNum.HasValue)
        {
            // If both have the same unit suffix (or none), compare directly
            var actualUnit = ExtractUnit(actual);
            var expectedUnit = ExtractUnit(expected);
            if (string.Equals(actualUnit, expectedUnit, StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrEmpty(actualUnit) || string.IsNullOrEmpty(expectedUnit))
            {
                return actualNum.Value.CompareTo(expectedNum.Value);
            }
        }

        // Try EMU-based dimension comparison (handles mixed units: "2cm" vs "1in")
        if (EmuConverter.TryParseEmu(actual, out var actualEmu) && EmuConverter.TryParseEmu(expected, out var expectedEmu))
        {
            return actualEmu.CompareTo(expectedEmu);
        }

        // Fallback: plain number comparison
        if (actualNum.HasValue && expectedNum.HasValue)
            return actualNum.Value.CompareTo(expectedNum.Value);

        // Last resort: string comparison
        return string.Compare(actual, expected, StringComparison.OrdinalIgnoreCase);
    }

    private static decimal? ExtractNumber(string value)
    {
        if (string.IsNullOrEmpty(value)) return null;

        // Strip known unit suffixes
        var trimmed = value.TrimEnd();
        foreach (var suffix in new[] { "pt", "px", "cm", "in", "em", "%" })
        {
            if (trimmed.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed[..^suffix.Length];
                break;
            }
        }

        return decimal.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    private static string ExtractUnit(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        foreach (var suffix in new[] { "pt", "px", "cm", "in", "em", "%" })
        {
            if (value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                return suffix;
        }
        return "";
    }
}
