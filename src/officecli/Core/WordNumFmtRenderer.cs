// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Converts a 1-based counter into the OOXML <c>w:numFmt</c> marker glyphs.
/// Covers the numFmt enum from ECMA-376 §17.18.59 that Word ships with;
/// unknown or unmapped values fall back to decimal.
/// </summary>
public static class WordNumFmtRenderer
{
    public static string Render(int n, string? numFmt)
    {
        if (n < 1) n = 1;
        switch ((numFmt ?? "decimal").ToLowerInvariant())
        {
            case "decimal": return n.ToString(CultureInfo.InvariantCulture);
            case "decimalzero": return n < 10 ? $"0{n}" : n.ToString(CultureInfo.InvariantCulture);
            case "upperroman": return ToRoman(n).ToUpperInvariant();
            case "lowerroman": return ToRoman(n).ToLowerInvariant();
            case "upperletter": return ToAlpha(n, uppercase: true);
            case "lowerletter": return ToAlpha(n, uppercase: false);
            case "ordinal": return ToOrdinal(n);
            case "cardinaltext": return ToEnglishCardinal(n);
            case "ordinaltext": return ToEnglishOrdinal(n);
            case "chinesecounting":
            case "japanesecounting":
                return ToChineseCounting(n, formal: false);
            case "chinesecountingthousand":
            case "taiwanesecounting":
            case "taiwanesecountingthousand":
                return ToChineseCounting(n, formal: true);
            case "chineselegalsimplified":
                return ToChineseLegalSimplified(n);
            case "ideographdigital":
            case "taiwanesedigital":
            case "japanesedigitaltenthousand":
                return ToIdeographDigital(n);
            case "koreandigital":
            case "koreandigital2":
                return ToKoreanDigital(n);
            case "koreancounting":
                return ToKoreanCounting(n);
            case "koreanlegal":
                return ToKoreanLegal(n);
            case "japaneselegal":
                return ToJapaneseLegal(n);
            case "ideographtraditional":
                return ToHeavenlyStems(n);
            case "ideographzodiac":
                return ToEarthlyBranches(n);
            case "decimalenclosedcircle":
            case "decimalenclosedcirclechinese":
                return ToEnclosedCircle(n);
            case "decimalenclosedfullstop":
                return $"{n}．";
            case "decimalenclosedparen":
                return $"({n})";
            case "decimalfullwidth":
            case "decimalfullwidth2":
                return ToFullWidthDigits(n);
            case "decimalhalfwidth":
                return n.ToString(CultureInfo.InvariantCulture);
            case "arabicabjad":
                return ToArabicAbjad(n);
            case "arabicalpha":
                return ToArabicAlpha(n);
            case "hebrew1":
            case "hebrew2":
                return ToHebrewNumeral(n);
            case "thainumbers":
            case "thaicounting":
                return ToThaiDigits(n);
            case "thailetters":
                return ToThaiLetters(n);
            case "hindinumbers":
            case "hindicounting":
            case "hindicardinaltext":
                return ToDevanagariDigits(n);
            case "hindiletters":
                return ToHindiLetters(n);
            case "hindivowels":
                return ToHindiVowels(n);
            case "russianlower":
                return ToRussianAlpha(n, uppercase: false);
            case "russianupper":
                return ToRussianAlpha(n, uppercase: true);
            case "none": return "";
            case "bullet": return "\u2022";
            default: return n.ToString(CultureInfo.InvariantCulture);
        }
    }

    // ---------- helpers ----------

    private static string ToRoman(int n)
    {
        if (n <= 0 || n > 3999) return n.ToString(CultureInfo.InvariantCulture);
        int[] vals = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] syms = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new StringBuilder();
        for (int i = 0; i < vals.Length; i++)
            while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
        return sb.ToString();
    }

    private static string ToAlpha(int n, bool uppercase)
    {
        // Word's behavior: A,B,...,Z,AA,BB,CC,... (repeating letter at 27+), not Excel column-style.
        var letter = (char)(((n - 1) % 26) + (uppercase ? 'A' : 'a'));
        var repeat = ((n - 1) / 26) + 1;
        return new string(letter, repeat);
    }

    private static string ToOrdinal(int n)
    {
        int mod100 = n % 100, mod10 = n % 10;
        string suffix = (mod100 is >= 11 and <= 13) ? "th" : mod10 switch
        {
            1 => "st", 2 => "nd", 3 => "rd", _ => "th"
        };
        return $"{n}{suffix}";
    }

    private static readonly string[] EnglishOnes =
    {
        "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
        "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
        "Seventeen", "Eighteen", "Nineteen"
    };
    private static readonly string[] EnglishTens =
    {
        "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"
    };

    private static string ToEnglishCardinal(int n)
    {
        if (n == 0) return "Zero";
        if (n < 0) return $"Negative {ToEnglishCardinal(-n)}";
        var sb = new StringBuilder();
        if (n >= 1000) { sb.Append(ToEnglishCardinal(n / 1000)).Append(" Thousand"); n %= 1000; if (n > 0) sb.Append(' '); }
        if (n >= 100) { sb.Append(EnglishOnes[n / 100]).Append(" Hundred"); n %= 100; if (n > 0) sb.Append(' '); }
        if (n >= 20) { sb.Append(EnglishTens[n / 10]); n %= 10; if (n > 0) sb.Append('-').Append(EnglishOnes[n]); }
        else if (n > 0) sb.Append(EnglishOnes[n]);
        return sb.ToString();
    }

    private static string ToEnglishOrdinal(int n)
    {
        var card = ToEnglishCardinal(n);
        // Only transform the trailing word.
        var lastSpace = card.LastIndexOf(' ');
        var lastHyphen = card.LastIndexOf('-');
        var split = Math.Max(lastSpace, lastHyphen);
        var head = split >= 0 ? card[..(split + 1)] : "";
        var tail = split >= 0 ? card[(split + 1)..] : card;
        string suffixMap(string w) => w switch
        {
            "One" => "First", "Two" => "Second", "Three" => "Third", "Five" => "Fifth",
            "Eight" => "Eighth", "Nine" => "Ninth", "Twelve" => "Twelfth",
            _ => w.EndsWith("y", StringComparison.Ordinal) ? w[..^1] + "ieth"
                 : w.EndsWith("e", StringComparison.Ordinal) ? w[..^1] + "th"
                 : w + "th"
        };
        return head + suffixMap(tail);
    }

    private static readonly char[] CnDigits = { '零', '一', '二', '三', '四', '五', '六', '七', '八', '九' };
    private static readonly char[] CnFormalDigits = { '零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖' };
    private static readonly char[] CnLegalSimplDigits = { '零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖' };

    private static string ToChineseCounting(int n, bool formal)
    {
        var digits = formal ? CnFormalDigits : CnDigits;
        char shi = formal ? '拾' : '十';
        char bai = formal ? '佰' : '百';
        char qian = formal ? '仟' : '千';
        char wan = formal ? '萬' : '万';
        return BuildCjkPositional(n, digits, shi, bai, qian, wan);
    }

    private static string ToChineseLegalSimplified(int n)
        => BuildCjkPositional(n, CnLegalSimplDigits, '拾', '佰', '仟', '万');

    private static string BuildCjkPositional(int n, char[] digits, char shi, char bai, char qian, char wan)
    {
        if (n == 0) return digits[0].ToString();
        if (n < 0) return "-" + BuildCjkPositional(-n, digits, shi, bai, qian, wan);
        if (n >= 10000)
        {
            var hi = n / 10000;
            var lo = n % 10000;
            var s = BuildCjkPositional(hi, digits, shi, bai, qian, wan) + wan;
            if (lo == 0) return s;
            if (lo < 1000) s += digits[0];
            return s + BuildCjkPositional(lo, digits, shi, bai, qian, wan);
        }
        // 0..9999
        var sb = new StringBuilder();
        int q = n / 1000, b = (n / 100) % 10, sh = (n / 10) % 10, u = n % 10;
        bool emittedNonZero = false;
        bool pendingZero = false;
        void emitDigit(int d, char? unit)
        {
            if (d == 0)
            {
                if (emittedNonZero) pendingZero = true;
                return;
            }
            if (pendingZero) { sb.Append(digits[0]); pendingZero = false; }
            // Special case: leading "一十" → "十" in informal spelling when n<20.
            if (unit == shi && d == 1 && !emittedNonZero)
                sb.Append(unit);
            else
            {
                sb.Append(digits[d]);
                if (unit.HasValue) sb.Append(unit.Value);
            }
            emittedNonZero = true;
        }
        emitDigit(q, qian);
        emitDigit(b, bai);
        emitDigit(sh, shi);
        emitDigit(u, null);
        return sb.ToString();
    }

    private static string ToIdeographDigital(int n)
    {
        // 〇一二三四五六七八九, positional: 25 → 二五, 100 → 一〇〇
        var s = n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
            sb.Append(c == '0' ? '〇' : CnDigits[c - '0']);
        return sb.ToString();
    }

    private static readonly string[] HeavenlyStems =
        { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
    private static readonly string[] EarthlyBranches =
        { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };

    private static string ToHeavenlyStems(int n) => HeavenlyStems[(n - 1) % 10];
    private static string ToEarthlyBranches(int n) => EarthlyBranches[(n - 1) % 12];

    private static string ToEnclosedCircle(int n)
    {
        // ① .. ⑳ = U+2460..U+2473 (1..20)
        if (n >= 1 && n <= 20) return ((char)(0x2460 + n - 1)).ToString();
        // 21..35 at U+3251..U+325F (Word uses similar enclosed glyphs); fallback to (n)
        if (n >= 21 && n <= 35) return ((char)(0x3251 + n - 21)).ToString();
        if (n >= 36 && n <= 50) return ((char)(0x32B1 + n - 36)).ToString();
        return $"({n})";
    }

    private static string ToFullWidthDigits(int n)
    {
        var s = n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
            sb.Append(c is >= '0' and <= '9' ? (char)('\uFF10' + (c - '0')) : c);
        return sb.ToString();
    }

    // Arabic alphabet (abjad order): 1..28
    private static readonly string[] AbjadLetters =
    {
        "أ", "ب", "ج", "د", "ه", "و", "ز", "ح", "ط", "ي",
        "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق", "ر",
        "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ"
    };
    private static string ToArabicAbjad(int n)
        => n >= 1 && n <= AbjadLetters.Length
            ? AbjadLetters[n - 1]
            : n.ToString(CultureInfo.InvariantCulture);

    // Arabic alphabet (alphabetical / hijā'ī order): 1..28
    private static readonly string[] ArabicAlphaLetters =
    {
        "أ", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر",
        "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ", "ف",
        "ق", "ك", "ل", "م", "ن", "ه", "و", "ي"
    };
    private static string ToArabicAlpha(int n)
        => n >= 1 && n <= ArabicAlphaLetters.Length
            ? ArabicAlphaLetters[n - 1]
            : n.ToString(CultureInfo.InvariantCulture);

    // Hebrew numerals (gematria), supports 1..999.
    private static string ToHebrewNumeral(int n)
    {
        if (n < 1 || n > 999) return n.ToString(CultureInfo.InvariantCulture);
        string[] ones = { "", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט" };
        string[] tens = { "", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ" };
        string[] hundreds = { "", "ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק" };
        var sb = new StringBuilder();
        sb.Append(hundreds[n / 100]);
        int rem = n % 100;
        if (rem == 15) sb.Append("טו");
        else if (rem == 16) sb.Append("טז");
        else { sb.Append(tens[rem / 10]); sb.Append(ones[rem % 10]); }
        return sb.ToString();
    }

    private static readonly string[] RussianAlphaLower =
    {
        "а", "б", "в", "г", "д", "е", "ж", "з", "и", "к",
        "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф",
        "х", "ц", "ч", "ш", "щ", "э", "ю", "я"
    };
    // Korean numerals ------------------------------------------------------

    private static readonly char[] KoreanSinoDigits = // 〇일이삼사오육칠팔구
        { '〇', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구' };
    private static readonly string[] KoreanNativeCounting = // 하나..열
        { "", "하나", "둘", "셋", "넷", "다섯", "여섯", "일곱", "여덟", "아홉", "열" };

    /// <summary>Positional sino-korean digits: 1 → 일, 25 → 이오, 100 → 일〇〇.</summary>
    private static string ToKoreanDigital(int n)
    {
        var s = n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
            sb.Append(c == '0' ? '〇' : KoreanSinoDigits[c - '0']);
        return sb.ToString();
    }

    /// <summary>Native Korean counting 1..10, beyond that falls back to sino-korean digital.</summary>
    private static string ToKoreanCounting(int n)
        => n is >= 1 and <= 10 ? KoreanNativeCounting[n] : ToKoreanDigital(n);

    /// <summary>Korean legal (formal) numerals share the Chinese formal hanzi set.</summary>
    private static string ToKoreanLegal(int n)
        => ToChineseCounting(n, formal: true);

    /// <summary>Japanese legal uses modern formal kanji 壱弐参肆伍陸漆捌玖拾.</summary>
    private static readonly char[] JpFormalDigits =
        { '零', '壱', '弐', '参', '肆', '伍', '陸', '漆', '捌', '玖' };
    private static string ToJapaneseLegal(int n)
        => BuildCjkPositional(n, JpFormalDigits, '拾', '佰', '仟', '萬');

    // Thai & Devanagari ----------------------------------------------------

    /// <summary>Positional Thai digits ๐๑๒...: 1 → ๑, 25 → ๒๕.</summary>
    private static string ToThaiDigits(int n)
    {
        var s = n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
            sb.Append(c is >= '0' and <= '9' ? (char)('\u0E50' + (c - '0')) : c);
        return sb.ToString();
    }

    // Thai consonants (44 letters), Word cycles after 44.
    private static string ToThaiLetters(int n)
    {
        // U+0E01..U+0E2E are the 46 code points but ฃ (U+0E03) and ฅ (U+0E05)
        // are obsolete; Word's enumeration skips them.
        char[] letters =
        {
            '\u0E01','\u0E02','\u0E04','\u0E06','\u0E07','\u0E08','\u0E09','\u0E0A','\u0E0B',
            '\u0E0C','\u0E0D','\u0E0E','\u0E0F','\u0E10','\u0E11','\u0E12','\u0E13','\u0E14',
            '\u0E15','\u0E16','\u0E17','\u0E18','\u0E19','\u0E1A','\u0E1B','\u0E1C','\u0E1D',
            '\u0E1E','\u0E1F','\u0E20','\u0E21','\u0E22','\u0E23','\u0E24','\u0E25','\u0E26',
            '\u0E27','\u0E28','\u0E29','\u0E2A','\u0E2B','\u0E2C','\u0E2D','\u0E2E'
        };
        return letters[(n - 1) % letters.Length].ToString();
    }

    /// <summary>Positional Devanagari digits ०१२...: 1 → १, 25 → २५.</summary>
    private static string ToDevanagariDigits(int n)
    {
        var s = n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
            sb.Append(c is >= '0' and <= '9' ? (char)('\u0966' + (c - '0')) : c);
        return sb.ToString();
    }

    // Devanagari consonants क, ख, ग, ...
    private static string ToHindiLetters(int n)
    {
        char[] letters =
        {
            'क','ख','ग','घ','ङ','च','छ','ज','झ','ञ',
            'ट','ठ','ड','ढ','ण','त','थ','द','ध','न',
            'प','फ','ब','भ','म','य','र','ल','व','श',
            'ष','स','ह'
        };
        return letters[(n - 1) % letters.Length].ToString();
    }

    // Devanagari vowels अ, आ, इ, ...
    private static string ToHindiVowels(int n)
    {
        char[] vowels = { 'अ','आ','इ','ई','उ','ऊ','ऋ','ए','ऐ','ओ','औ' };
        return vowels[(n - 1) % vowels.Length].ToString();
    }

    private static string ToRussianAlpha(int n, bool uppercase)
    {
        if (n < 1 || n > RussianAlphaLower.Length)
            return n.ToString(CultureInfo.InvariantCulture);
        var s = RussianAlphaLower[n - 1];
        return uppercase ? s.ToUpperInvariant() : s;
    }
}
