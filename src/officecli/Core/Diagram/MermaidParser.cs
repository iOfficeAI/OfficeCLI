// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Mermaid flowchart-subset parser: text → <see cref="DiagramGraph"/> (semantic IR).
///
/// Handles the common real-world syntax that shows up in the wild:
///   direction   flowchart TD|TB|LR|RL|BT
///   node shapes [rect] {diamond} (round) ([stadium]) ((circle)) {{hexagon}}
///               [/parallelogram/] [\trapezoid\] [(database)] [[subroutine]] &gt;flag]
///   edges       A --&gt; B --&gt; C  (chained), A ---|no arrow|, -.-&gt; ==&gt; --o --x &lt;--&gt;
///               A -- text --&gt; B  (mid-text label), A --&gt;|text| B  (pipe label)
///               A &amp; B --&gt; C &amp; D  (group expansion)
///   ignored     subgraph/end/direction/style/class/classDef/linkStyle/click, %% comments
///
/// Unknown tokens degrade to null (edge dropped) — the parser never throws.
/// </summary>
public static class MermaidParser
{
    // Node identifiers accept Unicode letters/digits, not just ASCII — mermaid
    // itself allows them, and CJK / accented ids (开始, 判断, café) are common in
    // non-English flowcharts. \p{L} = any letter, \p{N} = any digit. Without
    // this a fully-Chinese flowchart parses to zero nodes ("diagram has no
    // nodes"). The ASCII-only edge-id / class-attach patterns below are left as
    // they were — those name mermaid-internal constructs, not user labels.
    private const string Id = @"([\p{L}\p{N}_]+)";

    // node-shape wrappers, most-specific first
    private static readonly (FlowShape Shape, Regex Pat)[] ShapePats =
    {
        (FlowShape.Stadium,       new Regex(@"^" + Id + @"\(\[(.*)\]\)$")),
        (FlowShape.Subroutine,    new Regex(@"^" + Id + @"\[\[(.*)\]\]$")),
        (FlowShape.Database,      new Regex(@"^" + Id + @"\[\((.*)\)\]$")),
        (FlowShape.Circle,        new Regex(@"^" + Id + @"\(\((.*)\)\)$")),
        (FlowShape.Hexagon,       new Regex(@"^" + Id + @"\{\{(.*)\}\}$")),
        (FlowShape.Decision,      new Regex(@"^" + Id + @"\{(.*)\}$")),
        (FlowShape.Parallelogram, new Regex(@"^" + Id + @"\[/(.*)/\]$")),
        (FlowShape.Parallelogram, new Regex(@"^" + Id + @"\[\\(.*)\\\]$")),
        (FlowShape.Flag,          new Regex(@"^" + Id + @">(.*)\]$")),
        (FlowShape.Terminator,    new Regex(@"^" + Id + @"\((.*)\)$")),
        (FlowShape.Process,       new Regex(@"^" + Id + @"\[(.*)\]$")),
    };

    private static readonly Regex Bare = new(@"^" + Id + @"$");
    // link operator: optional leading '<', a run (>=2) of -/./=, optional head -/>/o/x, optional |label|
    private static readonly Regex Link = new(@"\s*(<?[-.=]{2,}[-.=]*[->oxX]?)\s*(?:\|([^|]*)\|)?\s*");
    // fold `A -- text --> B` (mid-text label) into pipe form `A -->|text| B`
    private static readonly Regex MidText = new(@"([-.=]{2,})\s+([^\-.=>|][^>|]*?)\s+([-.=]{2,}[->oxX])");
    private static readonly Regex Directive =
        new(@"^(subgraph|end|direction|click|style|classDef|class|linkStyle)\b", RegexOptions.IgnoreCase);
    private static readonly Regex DirHeader =
        new(@"^(?:flowchart|graph)\s+(TD|TB|LR|RL|BT)\b", RegexOptions.IgnoreCase);
    private static readonly Regex Header = new(@"^(?:flowchart|graph)\b", RegexOptions.IgnoreCase);
    private static readonly Regex EdgeId = new(@"^[A-Za-z0-9_]+@");
    // trailing class attach `A[label]:::className` — a style hook, not part of the id/shape
    private static readonly Regex ClassAttach = new(@":::[A-Za-z0-9_]+\s*$");

    public static DiagramGraph Parse(string text)
    {
        var g = new DiagramGraph();
        foreach (var line in Statements(text))
        {
            var s = line.Trim();
            if (s.Length == 0 || s.StartsWith("%%"))
                continue;

            var md = DirHeader.Match(s);
            if (md.Success)
            {
                var d = md.Groups[1].Value.ToUpperInvariant();
                g.Direction = (d == "LR" || d == "RL") ? FlowDirection.LeftRight : FlowDirection.TopDown;
                continue;
            }
            if (Header.IsMatch(s) || Directive.IsMatch(s))
                continue; // header / subgraph / style → no garbage nodes

            // Link operators are matched on a copy with node-label text masked
            // out, so `--` / `==` inside `A{count == 0}` or `A[Phase 1 -- draft]`
            // stays label text instead of splitting the node in two (which
            // dropped the node, its label and its edge). Indices are identical
            // in both strings, so every substring is taken from the original.
            var masked = MaskBracketText(s);
            var folded = new System.Text.StringBuilder();
            int at = 0;
            foreach (Match m in MidText.Matches(masked))
            {
                folded.Append(s, at, m.Index - at)
                      .Append(m.Groups[3].Value).Append('|')
                      .Append(s, m.Groups[2].Index, m.Groups[2].Length).Append('|');
                at = m.Index + m.Length;
            }
            if (at > 0)
            {
                s = folded.Append(s, at, s.Length - at).ToString();
                masked = MaskBracketText(s);
            }

            var links = Link.Matches(masked);
            if (links.Count == 0)
            {
                ParseNodeToken(s, g); // standalone node declaration
                continue;
            }

            var parts = new List<string>();
            var labels = new List<string>();
            int prev = 0;
            foreach (Match m in links)
            {
                parts.Add(s.Substring(prev, m.Index - prev));
                labels.Add(m.Groups[2].Success ? s.Substring(m.Groups[2].Index, m.Groups[2].Length) : "");
                prev = m.Index + m.Length;
            }
            parts.Add(s.Substring(prev));

            var groups = new List<List<string>>();
            foreach (var p in parts)
                groups.Add(Group(p, g));

            for (int i = 0; i < groups.Count - 1; i++)
            {
                var lbl = labels[i].Trim();
                foreach (var a in groups[i])
                    foreach (var b in groups[i + 1])
                        g.Edges.Add(new DiagramEdge { From = a, To = b, Label = lbl });
            }
        }
        return g;
    }

    private static IEnumerable<string> Statements(string text)
    {
        foreach (var raw in text.Split('\n'))
            foreach (var stmt in raw.Split(';'))
                yield return stmt;
    }

    /// <summary>Parse a single node token; register/update it. Returns id, or null if unparseable.</summary>
    private static string? ParseNodeToken(string tok, DiagramGraph g)
    {
        tok = tok.Trim();
        tok = EdgeId.Replace(tok, "").Trim();      // drop leading edge-id  (e1@--> )
        tok = ClassAttach.Replace(tok, "").Trim(); // drop trailing :::className (styling, ignored)
        if (tok.Length == 0)
            return null;

        foreach (var (shape, pat) in ShapePats)
        {
            var m = pat.Match(tok);
            if (!m.Success)
                continue;
            var id = m.Groups[1].Value;
            var lbl = m.Groups[2].Value.Trim().Trim('"').Trim('\'');
            if (lbl.Length == 0)
                lbl = id;
            var n = g.GetOrAdd(id);
            n.Label = lbl;
            n.Shape = shape;
            return id;
        }

        var mb = Bare.Match(tok);
        if (mb.Success)
        {
            var id = mb.Groups[1].Value;
            g.GetOrAdd(id);
            return id;
        }
        return null; // unparseable → skip, never throw
    }

    /// <summary>Expand an `A &amp; B` group into node ids (bracket-aware, so `&amp;` inside a label is safe).</summary>
    private static List<string> Group(string token, DiagramGraph g)
    {
        var ids = new List<string>();
        foreach (var t in SplitTop(token, '&'))
        {
            var id = ParseNodeToken(t, g);
            if (id != null)
                ids.Add(id);
        }
        return ids;
    }

    /// <summary>
    /// Same-length copy of <paramref name="s"/> with the text inside node-label
    /// brackets replaced by '_', so the link regexes only see depth-0 text
    /// (same bracket-depth rule as <see cref="SplitTop"/>). An unclosed
    /// bracket returns <paramref name="s"/> unchanged (the previous behaviour).
    /// </summary>
    private static string MaskBracketText(string s)
    {
        var chars = s.ToCharArray();
        int depth = 0;
        for (int i = 0; i < chars.Length; i++)
        {
            var ch = chars[i];
            if (ch is '[' or '(' or '{') { depth++; continue; }
            if (ch is ']' or ')' or '}') { depth = depth > 0 ? depth - 1 : 0; continue; }
            if (depth > 0) chars[i] = '_';
        }
        return depth == 0 ? new string(chars) : s;
    }

    /// <summary>Split on <paramref name="sep"/> only at bracket depth 0.</summary>
    private static List<string> SplitTop(string token, char sep)
    {
        var outp = new List<string>();
        int depth = 0;
        var cur = new System.Text.StringBuilder();
        foreach (var ch in token)
        {
            if (ch is '[' or '(' or '{') depth++;
            else if (ch is ']' or ')' or '}') depth = depth > 0 ? depth - 1 : 0;

            if (ch == sep && depth == 0)
            {
                outp.Add(cur.ToString());
                cur.Clear();
            }
            else
            {
                cur.Append(ch);
            }
        }
        outp.Add(cur.ToString());
        return outp;
    }
}
