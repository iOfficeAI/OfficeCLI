// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Text.RegularExpressions;

namespace OfficeCli;

/// <summary>
/// Loads help content from wiki markdown files.
/// Priority: local wiki directory (dev) > embedded resource (CI build) > null (fallback to hardcoded).
///
/// Three-layer help navigation:
///   officecli pptx set              → Layer 1: overview (all elements for this format+verb)
///   officecli pptx set shape        → Layer 2: element detail (one wiki page)
///   officecli pptx set shape.fill   → Layer 3: property detail (section from wiki page)
/// </summary>
internal static class WikiHelpLoader
{
    private static string? _localWikiDir;
    private static bool _localWikiChecked;

    // Format mapping: CLI format name → wiki prefix
    private static readonly Dictionary<string, string> FormatPrefix = new(StringComparer.OrdinalIgnoreCase)
    {
        ["docx"] = "word",
        ["xlsx"] = "excel",
        ["pptx"] = "ppt",
    };

    // Overview files per format
    private static readonly Dictionary<string, string> OverviewFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["docx"] = "word-reference.md",
        ["xlsx"] = "excel-reference.md",
        ["pptx"] = "powerpoint-reference.md",
    };

    // Command-level files: command-{verb}-{wikiFormat}.md
    private static readonly Dictionary<string, string> CommandFormatMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["docx"] = "word",
        ["xlsx"] = "excel",
        ["pptx"] = "ppt",
    };

    /// <summary>
    /// Try to load wiki-based help for the given format, verb, and optional element/property.
    /// Returns null if wiki content is not available (caller should fall back to hardcoded help).
    /// </summary>
    internal static string? TryGetHelp(string format, string? verb, string? element = null, string? property = null)
    {
        if (verb == null)
        {
            // Layer 1 overview: format reference page
            return TryLoadAndRender(OverviewFiles.GetValueOrDefault(format));
        }

        if (element == null)
        {
            // Layer 1 with verb: concatenate all element pages for this format+verb
            return TryLoadLayer1(format, verb);
        }

        // Layer 2 or 3: specific element page
        if (!FormatPrefix.TryGetValue(format, out var prefix)) return null;
        var fileName = $"{prefix}-{element}-{verb}.md";
        var content = TryLoadAndRender(fileName);
        if (content == null) return null;

        if (property != null)
        {
            // Layer 3: extract specific property section
            return ExtractPropertySection(content, property);
        }

        return content;
    }

    /// <summary>
    /// Layer 1 with verb: build a combined overview from all element pages.
    /// First shows the command-level page, then appends a summary of each element page.
    /// </summary>
    private static string? TryLoadLayer1(string format, string verb)
    {
        if (!CommandFormatMap.TryGetValue(format, out var cmdFormat)) return null;

        // Try command-level page first (e.g. command-set-ppt.md)
        var commandPage = TryLoadAndRender($"command-{verb}-{cmdFormat}.md");
        // If no command page, try generic (e.g. command-set.md)
        commandPage ??= TryLoadAndRender($"command-{verb}.md");

        if (!FormatPrefix.TryGetValue(format, out var prefix)) return commandPage;

        // Find all element pages for this format+verb
        var elementFiles = FindWikiFiles($"{prefix}-*-{verb}.md");
        if (elementFiles.Count == 0) return commandPage;

        var sb = new System.Text.StringBuilder();
        if (commandPage != null)
        {
            sb.AppendLine(commandPage);
            sb.AppendLine();
        }

        // Append each element page
        foreach (var file in elementFiles)
        {
            var content = TryLoadAndRender(file);
            if (content != null)
            {
                sb.AppendLine(content);
                sb.AppendLine();
            }
        }

        var result = sb.ToString().TrimEnd();
        return result.Length > 0 ? result : null;
    }

    /// <summary>
    /// Extract a property section from rendered help content.
    /// Looks for the property name in markdown table rows AND related sections.
    /// </summary>
    private static string? ExtractPropertySection(string content, string property)
    {
        var lines = content.Split('\n');
        var result = new System.Text.StringBuilder();
        var found = false;

        // Strategy 1: Find in markdown table rows (| `property` | ... |)
        for (int idx = 0; idx < lines.Length; idx++)
        {
            var line = lines[idx];
            if (line.Contains($"`{property}`", StringComparison.OrdinalIgnoreCase) ||
                line.Contains($"| {property} ", StringComparison.OrdinalIgnoreCase) ||
                line.Contains($"| `{property}` ", StringComparison.OrdinalIgnoreCase) ||
                line.Contains($"| {property} /", StringComparison.OrdinalIgnoreCase) ||
                line.Contains($"| `{property}` /", StringComparison.OrdinalIgnoreCase))
            {
                if (!found)
                {
                    // Include the table header if we can find it
                    for (int i = idx - 1; i >= 0; i--)
                    {
                        if (lines[i].TrimStart().StartsWith("| ") && lines[i].Contains("---"))
                        {
                            if (i > 0) result.AppendLine(lines[i - 1]);
                            result.AppendLine(lines[i]);
                            break;
                        }
                    }
                    found = true;
                }
                result.AppendLine(line);
            }
        }

        // Strategy 2: Also find any section heading that contains the property name
        // (e.g. "## Transition Format", "## Gradient Format" for property "transition"/"gradient")
        var escapedProp = Regex.Escape(property);
        var inSection = false;
        var sectionDepth = 0;
        for (int idx = 0; idx < lines.Length; idx++)
        {
            var line = lines[idx];
            if (!inSection)
            {
                var match = Regex.Match(line, @"^(#{1,4})\s+.*\b" + escapedProp + @"\b", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    inSection = true;
                    sectionDepth = match.Groups[1].Value.Length;
                    if (found) result.AppendLine(); // separator from table row
                    result.AppendLine(line);
                    found = true;
                    continue;
                }
            }
            else
            {
                // Stop at next heading of same or higher level
                var headingMatch = Regex.Match(line, @"^(#{1,4})\s+");
                if (headingMatch.Success && headingMatch.Groups[1].Value.Length <= sectionDepth)
                {
                    inSection = false;
                    // Don't break — there might be more matching sections
                    continue;
                }
                result.AppendLine(line);
            }
        }

        var sectionResult = result.ToString().TrimEnd();
        return sectionResult.Length > 0 ? sectionResult : null;
    }

    /// <summary>
    /// Try to load a wiki file and render it to terminal-friendly text.
    /// </summary>
    private static string? TryLoadAndRender(string? fileName)
    {
        if (fileName == null) return null;

        // 1. Try local wiki directory
        var localContent = TryLoadFromLocalWiki(fileName);
        if (localContent != null) return RenderMarkdown(localContent);

        // 2. Try embedded resource
        var embeddedContent = TryLoadFromEmbeddedResource(fileName);
        if (embeddedContent != null) return RenderMarkdown(embeddedContent);

        return null;
    }

    private static string? TryLoadFromLocalWiki(string fileName)
    {
        var dir = FindLocalWikiDir();
        if (dir == null) return null;
        var path = Path.Combine(dir, fileName);
        return File.Exists(path) ? File.ReadAllText(path) : null;
    }

    private static string? TryLoadFromEmbeddedResource(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = assembly.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith("." + fileName.Replace("/", ".").Replace("-", "_"),
                StringComparison.OrdinalIgnoreCase));

        // Also try with original dashes (resource names may vary)
        resourceName ??= assembly.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));

        if (resourceName == null) return null;

        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Find local wiki directory by looking for a sibling OfficeCli.wiki directory
    /// relative to the git repository root.
    /// </summary>
    private static string? FindLocalWikiDir()
    {
        if (_localWikiChecked) return _localWikiDir;
        _localWikiChecked = true;

        try
        {
            var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (dir != null)
            {
                if (Directory.Exists(Path.Combine(dir.FullName, ".git")))
                {
                    // Found git root, look for sibling wiki directory
                    var wikiDir = Path.Combine(dir.Parent?.FullName ?? dir.FullName, "OfficeCli.wiki");
                    if (Directory.Exists(wikiDir))
                    {
                        _localWikiDir = wikiDir;
                        return _localWikiDir;
                    }
                    break;
                }
                dir = dir.Parent;
            }
        }
        catch
        {
            // Ignore filesystem errors during discovery
        }

        return null;
    }

    /// <summary>
    /// Find wiki files matching a glob pattern from local dir or embedded resources.
    /// Returns file names (not full paths).
    /// </summary>
    private static List<string> FindWikiFiles(string pattern)
    {
        var files = new List<string>();

        // Try local wiki dir
        var dir = FindLocalWikiDir();
        if (dir != null)
        {
            try
            {
                foreach (var path in Directory.GetFiles(dir, pattern))
                    files.Add(Path.GetFileName(path));
            }
            catch { }
        }

        // If no local files, try embedded resources
        if (files.Count == 0)
        {
            var assembly = Assembly.GetExecutingAssembly();
            // Convert glob pattern to regex for matching resource names
            var regexPattern = "\\." + Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("-", "[_\\-]") + "$";
            foreach (var name in assembly.GetManifestResourceNames())
            {
                if (Regex.IsMatch(name, regexPattern, RegexOptions.IgnoreCase))
                {
                    // Extract original file name from resource name
                    var parts = name.Split('.');
                    if (parts.Length >= 2)
                    {
                        // Resource name format: Namespace.wiki.file_name.md
                        var fn = string.Join(".", parts.Skip(parts.Length - 2));
                        files.Add(fn);
                    }
                }
            }
        }

        files.Sort(StringComparer.OrdinalIgnoreCase);
        return files;
    }

    /// <summary>
    /// Render markdown to terminal-friendly text.
    /// Keeps tables and code blocks as-is (they're already text-friendly).
    /// </summary>
    internal static string RenderMarkdown(string markdown)
    {
        var lines = markdown.Split('\n');
        var sb = new System.Text.StringBuilder();

        foreach (var line in lines)
        {
            var rendered = line;

            // Remove footer lines like "---" and "*Based on OfficeCLI v1.0.11*"
            if (rendered.TrimStart() == "---" || rendered.TrimStart().StartsWith("*Based on"))
                continue;

            // Strip markdown link syntax [text](url) → text
            rendered = Regex.Replace(rendered, @"\[([^\]]+)\]\([^)]+\)", "$1");

            // Strip inline code backticks in table cells for cleaner output
            // But keep code blocks (``` lines) intact
            if (!rendered.TrimStart().StartsWith("```"))
                rendered = rendered.Replace("`", "");

            sb.AppendLine(rendered);
        }

        return sb.ToString().TrimEnd();
    }
}
