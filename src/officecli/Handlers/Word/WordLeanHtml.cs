// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Handlers;

/// <summary>
/// Lean static profile for the Word HTML preview: one pass over the finished
/// page that drops watch/goto/range-screenshot scaffolding and hoists inline
/// style attributes into a generated class sheet. The page's own &lt;style&gt;
/// blocks are wrapped in <c>@layer base</c> and the generated rules stay
/// unlayered, so a hoisted rule beats every base rule regardless of
/// specificity, exactly as the inline style did (CSS Cascade 5). Important
/// base declarations still win, as they did over inline styles.
/// <para>Declarations the preview script or base CSS read from the style
/// attribute (float, column-count, the .page element) stay inline. Script,
/// style, title and textarea contents are copied verbatim, so header/footer
/// templates embedded as JS strings keep their inline styles.</para>
/// </summary>
internal static class WordLeanHtml
{
    internal sealed record Options
    {
        public bool StripMarkers { get; init; } = true;
        public bool StripAnchors { get; init; } = true;
        public bool StripDataPath { get; init; } = true;
        public bool StripColTwips { get; init; } = true;
        public bool InternStyles { get; init; } = true;
    }

    private const StringComparison OIC = StringComparison.OrdinalIgnoreCase;
    private static readonly Regex MarkerRx = new(@"<span class=""w[be]"" data-block=""\d+"" style=""display:none""></span>", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex AnchorRx = new(@"<a id=""w-(?:p|table)-\d+""></a>", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex ColTwipsDeclRx = new(@"(?:^|;)\s*--col-twips\s*:\s*\d+\s*(?=;|$)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex FloatRx = new(@"float\s*:", RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
    private static readonly string[] RawTextElements = { "script", "style", "title", "textarea" };
    private const int StyleElement = 1;
    private static readonly char[] SheetUnsafe = { '<', '{', '}' };

    private readonly record struct Attr(string Name, int Start, int End, string? Value);

    private sealed class Tag
    {
        public string Name = "";
        public int Start, NameEnd, TailStart, End;
        public bool SelfClosing;
        public readonly List<Attr> Attrs = new();
    }

    public static string Transform(string html, Options? options = null)
    {
        var o = options ?? new Options();
        if (o.StripMarkers) html = MarkerRx.Replace(html, "");
        if (o.StripAnchors) html = AnchorRx.Replace(html, "");
        if (!o.StripDataPath && !o.StripColTwips && !o.InternStyles) return html;

        var classes = new Dictionary<string, string>(StringComparer.Ordinal);
        var sheet = new StringBuilder();
        var sb = new StringBuilder(html.Length);
        int headCloseAt = -1, i = 0, n = html.Length;
        while (i < n)
        {
            int lt = html.IndexOf('<', i);
            if (lt < 0) { sb.Append(html, i, n - i); break; }
            sb.Append(html, i, lt - i);
            if (string.CompareOrdinal(html, lt, "<!--", 0, 4) == 0)
            {
                int end = html.IndexOf("-->", lt + 4, StringComparison.Ordinal);
                end = end < 0 ? n : end + 3;
                sb.Append(html, lt, end - lt); i = end; continue;
            }
            char c1 = lt + 1 < n ? html[lt + 1] : '\0';
            if (c1 is '/' or '!' or '?')
            {
                int gt = html.IndexOf('>', lt);
                int end = gt < 0 ? n : gt + 1;
                if (headCloseAt < 0 && c1 == '/' && string.Compare(html, lt + 2, "head", 0, 4, OIC) == 0) headCloseAt = sb.Length;
                sb.Append(html, lt, end - lt); i = end; continue;
            }
            if (!char.IsAsciiLetter(c1)) { sb.Append('<'); i = lt + 1; continue; }
            if (!TryParseStartTag(html, lt, out var tag)) { sb.Append(html, lt, n - lt); break; }

            sb.Append(RewriteStartTag(html, tag, o, classes, sheet));
            i = tag.End;

            int raw = Array.FindIndex(RawTextElements, r => r.Equals(tag.Name, OIC));
            if (raw >= 0 && !tag.SelfClosing)
            {
                int close = html.IndexOf("</" + RawTextElements[raw], i, OIC);
                if (close < 0) close = n;
                var content = html.AsSpan(i, close - i);
                if (o.InternStyles && raw == StyleElement && !content.IsWhiteSpace())
                    sb.Append("@layer base {\n").Append(content).Append("\n}\n");
                else
                    sb.Append(content);
                i = close;
            }
        }
        if (sheet.Length > 0)
            sb.Insert(headCloseAt < 0 ? 0 : headCloseAt, "<style>\n" + sheet + "</style>\n");
        return sb.ToString();
    }

    private static bool TryParseStartTag(string h, int lt, out Tag tag)
    {
        tag = new Tag { Start = lt };
        int n = h.Length, k = lt + 1;
        while (k < n && (char.IsAsciiLetterOrDigit(h[k]) || h[k] is '-' or ':')) k++;
        tag.Name = h.Substring(lt + 1, k - lt - 1);
        tag.NameEnd = k;
        while (true)
        {
            int attrStart = k;
            while (k < n && char.IsWhiteSpace(h[k])) k++;
            if (k >= n) return false;
            if (h[k] == '>') { tag.TailStart = attrStart; tag.End = k + 1; return true; }
            if (h[k] == '/' && k + 1 < n && h[k + 1] == '>') { tag.TailStart = attrStart; tag.SelfClosing = true; tag.End = k + 2; return true; }
            int nameStart = k;
            while (k < n && !char.IsWhiteSpace(h[k]) && h[k] != '=' && h[k] != '>' && !(h[k] == '/' && k + 1 < n && h[k + 1] == '>')) k++;
            if (k == nameStart) k++;
            var name = h.Substring(nameStart, k - nameStart);
            int afterName = k;
            while (k < n && char.IsWhiteSpace(h[k])) k++;
            string? value = null;
            if (k < n && h[k] == '=')
            {
                k++;
                while (k < n && char.IsWhiteSpace(h[k])) k++;
                if (k >= n) return false;
                char q = h[k];
                if (q is '"' or '\'')
                {
                    int close = h.IndexOf(q, k + 1);
                    if (close < 0) return false;
                    value = h.Substring(k + 1, close - k - 1);
                    k = close + 1;
                }
                else
                {
                    int vs = k;
                    while (k < n && !char.IsWhiteSpace(h[k]) && h[k] != '>') k++;
                    value = h.Substring(vs, k - vs);
                }
            }
            else k = afterName;
            tag.Attrs.Add(new Attr(name, attrStart, k, value));
        }
    }

    private static bool IsDropped(string name, Options o) =>
        (o.StripDataPath && name.Equals("data-path", OIC)) || (o.StripColTwips && name.Equals("data-col-twips", OIC));

    private static bool IsPinned(string rawStyle, string? classValue) =>
        FloatRx.IsMatch(rawStyle)
        || rawStyle.Contains("column-count", OIC)
        || (classValue != null && classValue.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Contains("page"));

    private static string RewriteStartTag(string h, Tag t, Options o, Dictionary<string, string> classes, StringBuilder sheet)
    {
        int styleIdx = -1, classIdx = -1;
        bool unsafeTag = false, touch = false;
        for (int a = 0; a < t.Attrs.Count; a++)
        {
            var at = t.Attrs[a];
            if (at.Name.Equals("style", OIC)) { if (styleIdx >= 0) unsafeTag = true; else styleIdx = a; }
            else if (at.Name.Equals("class", OIC)) { if (classIdx >= 0) unsafeTag = true; else classIdx = a; }
            else if (IsDropped(at.Name, o)) touch = true;
            if ((at.Name.Equals("style", OIC) || at.Name.Equals("class", OIC)) && at.Value?.Contains('"') == true) unsafeTag = true;
        }
        if (unsafeTag) styleIdx = classIdx = -1;

        string? rewrittenStyle = null;
        if (styleIdx >= 0 && o.StripColTwips)
        {
            var v = t.Attrs[styleIdx].Value ?? "";
            if (v.Contains("--col-twips", StringComparison.Ordinal))
            { rewrittenStyle = ColTwipsDeclRx.Replace(v, "").TrimStart(';'); touch = true; }
        }

        string? internClass = null;
        if (o.InternStyles && styleIdx >= 0)
        {
            var raw = rewrittenStyle ?? t.Attrs[styleIdx].Value ?? "";
            if (raw.Length > 0 && !IsPinned(raw, classIdx >= 0 ? t.Attrs[classIdx].Value : null))
            {
                var css = WebUtility.HtmlDecode(raw);
                if (css.IndexOfAny(SheetUnsafe) < 0)
                {
                    if (!classes.TryGetValue(css, out internClass))
                    {
                        internClass = "s" + (classes.Count + 1);
                        classes[css] = internClass;
                        sheet.Append('.').Append(internClass).Append('{').Append(css).Append("}\n");
                    }
                    touch = true;
                }
            }
        }
        if (!touch) return h.Substring(t.Start, t.End - t.Start);

        var sb = new StringBuilder(t.End - t.Start);
        sb.Append(h, t.Start, t.NameEnd - t.Start);
        for (int a = 0; a < t.Attrs.Count; a++)
        {
            var at = t.Attrs[a];
            if (IsDropped(at.Name, o)) continue;
            if (a == styleIdx && internClass != null)
            {
                if (classIdx < 0) sb.Append(" class=\"").Append(internClass).Append('"');
                continue;
            }
            if (a == styleIdx && rewrittenStyle != null)
            {
                if (rewrittenStyle.Length > 0) sb.Append(" style=\"").Append(rewrittenStyle).Append('"');
                continue;
            }
            if (a == classIdx && internClass != null)
            {
                var cv = at.Value ?? "";
                sb.Append(" class=\"").Append(cv.Length > 0 ? cv + " " + internClass : internClass).Append('"');
                continue;
            }
            sb.Append(h, at.Start, at.End - at.Start);
        }
        sb.Append(h, t.TailStart, t.End - t.TailStart);
        return sb.ToString();
    }
}
