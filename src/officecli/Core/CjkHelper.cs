// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Modified by cli-jaw contributors
// Added: CJK font handling, language detection, kinsoku processing

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// CJK (Chinese-Japanese-Korean) text handling utilities.
/// Provides font fallback chains, language detection, and line-break rules.
/// Uses <see cref="ParseHelpers.IsCjkOrFullWidth"/> for basic character classification.
/// </summary>
public static class CjkHelper
{
    // ── Script detection ────────────────────────────────────────────

    /// <summary>Detect if text contains any CJK characters.</summary>
    public static bool ContainsCjk(string? text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (var c in text)
            if (ParseHelpers.IsCjkOrFullWidth(c))
                return true;
        return false;
    }

    /// <summary>Detect the dominant CJK script in text.</summary>
    public static CjkScript DetectScript(string? text)
    {
        if (string.IsNullOrEmpty(text)) return CjkScript.None;

        int ko = 0, ja = 0, zh = 0;
        foreach (var c in text)
        {
            if (IsKorean(c)) ko++;
            else if (IsJapanese(c)) ja++;
            else if (IsChinese(c)) zh++;
        }

        if (ko == 0 && ja == 0 && zh == 0) return CjkScript.None;
        if (ko >= ja && ko >= zh) return CjkScript.Korean;
        if (ja >= ko && ja >= zh) return CjkScript.Japanese;
        return CjkScript.Chinese;
    }

    // ── Font chains ─────────────────────────────────────────────────

    /// <summary>Get the font fallback chain for a CJK script.</summary>
    public static (string primary, string fallback1, string fallback2) GetFontChain(CjkScript script) =>
        script switch
        {
            CjkScript.Korean   => ("Malgun Gothic", "맑은 고딕", "AppleSDGothicNeo-Regular"),
            CjkScript.Japanese => ("Yu Gothic", "Meiryo", "Hiragino Sans"),
            CjkScript.Chinese  => ("Microsoft YaHei", "SimSun", "PingFang SC"),
            _ => ("", "", "")
        };

    /// <summary>Get the BCP 47 language tag for a CJK script.</summary>
    public static string GetLanguageTag(CjkScript script) =>
        script switch
        {
            CjkScript.Korean   => "ko-KR",
            CjkScript.Japanese => "ja-JP",
            CjkScript.Chinese  => "zh-CN",
            _ => ""
        };

    /// <summary>
    /// Split text into contiguous CJK/non-CJK segments so mixed content can be
    /// emitted as separate runs.
    /// </summary>
    public static IReadOnlyList<(string text, CjkScript script)> SegmentText(string? text)
    {
        var segments = new List<(string text, CjkScript script)>();
        if (string.IsNullOrEmpty(text)) return segments;

        var buffer = new StringBuilder();
        var currentIsCjk = ParseHelpers.IsCjkOrFullWidth(text[0]);

        foreach (var c in text)
        {
            var isCjk = ParseHelpers.IsCjkOrFullWidth(c);
            if (buffer.Length > 0 && isCjk != currentIsCjk)
            {
                AddSegment(segments, buffer.ToString(), currentIsCjk);
                buffer.Clear();
            }

            buffer.Append(c);
            currentIsCjk = isCjk;
        }

        if (buffer.Length > 0)
            AddSegment(segments, buffer.ToString(), currentIsCjk);

        return segments;
    }

    // ── WordML (DOCX) helpers ───────────────────────────────────────

    /// <summary>
    /// Apply CJK fonts and language to a WordML <c>w:rPr</c> element.
    /// Sets <c>w:rFonts/@w:eastAsia</c> and <c>w:lang/@w:eastAsia</c>.
    /// </summary>
    public static void ApplyToWordRun(RunProperties rPr, CjkScript script)
    {
        if (script == CjkScript.None || rPr == null) return;
        var (primary, _, _) = GetFontChain(script);
        var lang = GetLanguageTag(script);

        // w:rFonts — set eastAsia only (preserve user's Ascii/HighAnsi)
        var rFonts = rPr.GetFirstChild<RunFonts>();
        if (rFonts == null)
        {
            rFonts = new RunFonts();
            rPr.PrependChild(rFonts);
        }
        rFonts.EastAsia = primary;

        // w:lang — set eastAsia
        var langElem = rPr.GetFirstChild<Languages>();
        if (langElem == null)
        {
            langElem = new Languages();
            rPr.AppendChild(langElem);
        }
        langElem.EastAsia = lang;
    }

    /// <summary>Remove CJK-only WordML font and language metadata from a run.</summary>
    public static void ClearWordRunCjk(RunProperties rPr)
    {
        if (rPr == null) return;

        var rFonts = rPr.GetFirstChild<RunFonts>();
        if (rFonts != null)
            rFonts.EastAsia = null;

        var langElem = rPr.GetFirstChild<Languages>();
        if (langElem != null)
            langElem.EastAsia = null;
    }

    /// <summary>
    /// Detect CJK in text and apply font/lang to the run's properties.
    /// Creates <c>RunProperties</c> if missing.
    /// </summary>
    public static void ApplyToWordRunIfCjk(Run run, string? text)
    {
        var script = DetectScript(text);
        if (script == CjkScript.None)
        {
            var existingRPr = run.GetFirstChild<RunProperties>();
            if (existingRPr != null)
                ClearWordRunCjk(existingRPr);
            return;
        }

        var rPr = run.GetFirstChild<RunProperties>();
        if (rPr == null)
        {
            rPr = new RunProperties();
            run.PrependChild(rPr);
        }
        ApplyToWordRun(rPr, script);
    }

    // ── DrawingML (PPTX/XLSX chart text) helpers ────────────────────

    /// <summary>
    /// Apply CJK fonts and language to a DrawingML <c>a:rPr</c> element.
    /// Sets <c>a:ea/@typeface</c> and <c>a:rPr/@lang</c>.
    /// </summary>
    public static void ApplyToDrawingRun(A.RunProperties rPr, CjkScript script)
    {
        if (script == CjkScript.None || rPr == null) return;
        var (primary, _, _) = GetFontChain(script);
        var lang = GetLanguageTag(script);

        // a:ea (East Asian font)
        var eaFont = rPr.GetFirstChild<A.EastAsianFont>();
        if (eaFont == null)
        {
            eaFont = new A.EastAsianFont();
            rPr.AppendChild(eaFont);
        }
        eaFont.Typeface = primary;

        // lang attribute
        rPr.Language = lang;
    }

    /// <summary>Remove DrawingML CJK font metadata and restore fallback language.</summary>
    public static void ClearDrawingRunCjk(A.RunProperties rPr, string fallbackLang = "en-US")
    {
        if (rPr == null) return;

        rPr.RemoveAllChildren<A.EastAsianFont>();
        rPr.Language = fallbackLang;
    }

    /// <summary>
    /// Detect CJK in text and apply font/lang. Falls back to the provided
    /// default language when no CJK is detected.
    /// </summary>
    public static void ApplyToDrawingRunIfCjk(A.RunProperties rPr, string? text, string fallbackLang = "en-US")
    {
        var script = DetectScript(text);
        if (script != CjkScript.None)
            ApplyToDrawingRun(rPr, script);
        else
            ClearDrawingRunCjk(rPr, fallbackLang);
    }

    // ── Kinsoku (line-break rules) ──────────────────────────────────

    // Characters that must NOT appear at the start of a line
    private const string KinsokuStartChars =
        "!%),.:;?]}¢°·'\"†‡›℃∵、。〉》」』】〕〗〙〛！＂％＇），．：；？＞］｝～";

    // Characters that must NOT appear at the end of a line
    private const string KinsokuEndChars =
        "$(£¥·'\"〈《「『【〔〖〘〚＄（［｛￡￥";

    /// <summary>Cannot start a line (e.g. closing brackets, periods).</summary>
    public static bool IsKinsokuStart(char c) => KinsokuStartChars.Contains(c);

    /// <summary>Cannot end a line (e.g. opening brackets).</summary>
    public static bool IsKinsokuEnd(char c) => KinsokuEndChars.Contains(c);

    // ── Private character classifiers ───────────────────────────────

    private static bool IsKorean(char c) =>
        (c >= 0xAC00 && c <= 0xD7AF)    // Hangul Syllables
        || (c >= 0x1100 && c <= 0x11FF)  // Hangul Jamo
        || (c >= 0x3130 && c <= 0x318F); // Hangul Compat Jamo

    private static bool IsJapanese(char c) =>
        (c >= 0x3040 && c <= 0x309F)     // Hiragana
        || (c >= 0x30A0 && c <= 0x30FF)  // Katakana
        || (c >= 0x31F0 && c <= 0x31FF); // Katakana Phonetic Ext

    private static bool IsChinese(char c) =>
        (c >= 0x4E00 && c <= 0x9FFF)     // CJK Unified Ideographs
        || (c >= 0x3400 && c <= 0x4DBF); // CJK Extension A

    private static void AddSegment(List<(string text, CjkScript script)> segments, string text, bool isCjk)
    {
        segments.Add((text, isCjk ? DetectScript(text) : CjkScript.None));
    }
}

/// <summary>Identified CJK script family.</summary>
public enum CjkScript { None, Korean, Japanese, Chinese }
