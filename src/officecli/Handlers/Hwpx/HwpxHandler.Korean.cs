// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;

namespace OfficeCli.Handlers;

internal static partial class HwpxKorean
{
    public static string Normalize(string text)
    {
        text = StripPuaChars(text);
        text = StripShapeAltText(text);
        text = NormalizeKoreanSpacing(text);
        return text;
    }

    // F1: Rune-based PUA stripping — handles BMP + supplementary PUA planes
    public static string StripPuaChars(string text)
        => string.Concat(text.EnumerateRunes()
            .Where(r => !(r.Value >= 0xE000 && r.Value <= 0xF8FF)      // BMP PUA
                      && !(r.Value >= 0xF0000 && r.Value <= 0xFFFFD)    // PUA-A
                      && !(r.Value >= 0x100000 && r.Value <= 0x10FFFD)) // PUA-B
            .Select(r => r.ToString()));

    public static string StripShapeAltText(string text)
        => ShapeAltTextRegex().Replace(text, "");

    public static string NormalizeKoreanSpacing(string text)
    {
        // Fix uniform-distribution spacing (균등 분할): "현 장 대 응" → "현장대응"
        // Only collapse when 3+ consecutive single Korean syllables are space-separated.
        // Preserves normal word spacing like "인사 발령 통보".
        text = UniformDistRegex().Replace(text, m => m.Value.Replace(" ", ""));
        // Remove zero-width joiners between jamo
        text = text.Replace("\u200D", "");
        return text;
    }

    // Plan 99.9.A4: Complete shape alt-text regex (kordoc TB2, 50+ shapes)
    // Anchored with ^...$ to prevent partial-match false positives.
    [GeneratedRegex(@"^(?:모서리가 둥근 |둥근 )?(?:표|그림|개체|사각형|직사각형|정사각형|원|타원|삼각형|이등변 삼각형|직각 삼각형|선|직선|곡선|화살표|굵은 화살표|이중 화살표|오각형|육각형|팔각형|별|[4-8]점별|십자|십자형|구름|구름형|마름모|도넛|평행사변형|사다리꼴|부채꼴|호|반원|물결|번개|하트|빗금|블록 화살표|수식|그리기\s*개체|묶음\s*개체|글상자|수식\s*개체|OLE\s*개체)\s*입니다\.?$")]
    private static partial Regex ShapeAltTextRegex();

    // F2: Hangul Syllables + Compatibility Jamo uniform spacing detection
    [GeneratedRegex(@"(?<![\uAC00-\uD7A3\u3131-\u318E])[\uAC00-\uD7A3\u3131-\u318E](?: [\uAC00-\uD7A3\u3131-\u318E]){2,}(?![\uAC00-\uD7A3\u3131-\u318E])")]
    private static partial Regex UniformDistRegex();
}
