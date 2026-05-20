// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Shared C-style escape resolution for user-supplied text values
/// across docx/xlsx/pptx Add and Set paths.
///
/// Handlers historically called <c>value.Replace("\\n", "\n").Replace("\\t", "\t")</c>
/// inline. That naive two-pass form had no way to express a literal
/// backslash followed by 'n' / 't' — the user could not type "\\n" and
/// get the two-character string <c>\n</c> back, because the trailing
/// <c>\n</c> was always consumed by the second replace.
///
/// <see cref="Resolve"/> does a single left-to-right scan that recognizes
/// <c>\\</c> (literal backslash), <c>\n</c> (LF), <c>\t</c> (TAB), and
/// <c>\r</c> (CR). Unknown escape sequences are passed through verbatim
/// so today's behavior for stray backslashes (e.g. Windows paths typed
/// without doubling) doesn't regress.
/// </summary>
public static class TextEscape
{
    /// <summary>
    /// Resolve C-style escape sequences in <paramref name="value"/>.
    /// Returns the input unchanged when it contains no backslash.
    /// </summary>
    public static string Resolve(string? value)
    {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
        if (value.IndexOf('\\') < 0) return value;

        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            var ch = value[i];
            if (ch != '\\' || i + 1 >= value.Length)
            {
                sb.Append(ch);
                continue;
            }
            var next = value[i + 1];
            switch (next)
            {
                case '\\': sb.Append('\\'); i++; break;
                case 'n':  sb.Append('\n'); i++; break;
                case 't':  sb.Append('\t'); i++; break;
                case 'r':  sb.Append('\r'); i++; break;
                default:
                    // Unknown escape — pass the backslash through verbatim
                    // (subsequent char handled on next iteration).
                    sb.Append('\\');
                    break;
            }
        }
        return sb.ToString();
    }
}
