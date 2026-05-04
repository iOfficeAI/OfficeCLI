// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using OfficeCli.Core;
using OfficeCli.Handlers.Hwp.SafeSave;

namespace OfficeCli.Handlers.Hwpx.Validation;

internal static class HwpxPackageValidator
{
    public static HwpxPackageValidationResult Validate(string path, bool strictOrphans = false)
    {
        var checks = new List<SafeSaveCheck>();
        var package = new Dictionary<string, object?>(StringComparer.Ordinal);
        List<ValidationError> validationErrors;

        try
        {
            using var zip = ZipFile.OpenRead(path);
            var entryNames = zip.Entries.Select(entry => entry.FullName).ToArray();
            package["entryCount"] = entryNames.Length;
            package["xmlPartCount"] = entryNames.Count(IsXmlLikePart);
            checks.Add(new SafeSaveCheck("zip-open", true, "info", null, new Dictionary<string, object?>
            {
                ["entryCount"] = entryNames.Length
            }));
            checks.Add(PresenceCheck("manifest-present", HasEntry(entryNames, "Contents/content.hpf"), "Contents/content.hpf"));
            checks.Add(PresenceCheck("content-parts-present", entryNames.Any(IsSectionPart), "Contents/section*.xml"));
            checks.Add(PresenceCheck("header-parts-present", HasEntry(entryNames, "Contents/header.xml"), "Contents/header.xml"));
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
        {
            package["entryCount"] = 0;
            package["xmlPartCount"] = 0;
            package["error"] = ex.Message;
            checks.Add(new SafeSaveCheck("zip-open", false, "error", ex.Message));
            checks.Add(new SafeSaveCheck("package-integrity", false, "error", "HWPX ZIP package could not be opened."));
            return new HwpxPackageValidationResult(checks, package);
        }

        try
        {
            using var handler = new HwpxHandler(path, editable: false);
            validationErrors = handler.Validate();
        }
        catch (Exception ex)
        {
            var errorType = ex is System.Xml.XmlException
                ? "xml_malformed"
                : "hwpx_load_failed";
            validationErrors = [new ValidationError(errorType, ex.Message, "/", null)];
        }

        var xmlOk = !validationErrors.Any(error => ContainsAny(error.ErrorType, "xml", "malformed"));
        checks.Add(new SafeSaveCheck(
            "xml-well-formed",
            xmlOk,
            xmlOk ? "info" : "error",
            xmlOk ? null : "One or more HWPX XML parts are malformed."));

        var missingBinData = validationErrors.Count(error => error.ErrorType.Contains("bindata_missing", StringComparison.OrdinalIgnoreCase));
        var orphanBinData = validationErrors.Count(error => error.ErrorType.Contains("bindata_orphan", StringComparison.OrdinalIgnoreCase));
        checks.Add(new SafeSaveCheck(
            "bindata-references-present",
            missingBinData == 0,
            missingBinData == 0 ? "info" : "error",
            missingBinData == 0 ? null : "One or more BinData references point to missing package entries.",
            new Dictionary<string, object?> { ["missingCount"] = missingBinData }));
        checks.Add(new SafeSaveCheck(
            "orphan-reference-report",
            true,
            orphanBinData == 0 ? "info" : "warning",
            orphanBinData == 0 ? null : "Package contains unreferenced BinData entries.",
            new Dictionary<string, object?> { ["orphanCount"] = orphanBinData }));

        var blockingErrors = validationErrors
            .Where(error => IsPackageBlocking(error, strictOrphans))
            .ToArray();
        var packageOk = checks.Where(check => check.Name != "orphan-reference-report").All(check => check.Ok)
            && blockingErrors.Length == 0;

        package["validationErrorCount"] = validationErrors.Count;
        package["blockingErrorCount"] = blockingErrors.Length;
        package["missingBinDataCount"] = missingBinData;
        package["orphanBinDataCount"] = orphanBinData;
        package["strictOrphans"] = strictOrphans;
        checks.Add(new SafeSaveCheck(
            "package-integrity",
            packageOk,
            packageOk ? "info" : "error",
            packageOk ? null : "HWPX package integrity validation failed.",
            package));

        return new HwpxPackageValidationResult(checks, package);
    }

    private static SafeSaveCheck PresenceCheck(string name, bool ok, string expectedPath) => new(
        name,
        ok,
        ok ? "info" : "error",
        ok ? null : $"Missing required HWPX package part: {expectedPath}",
        new Dictionary<string, object?> { ["expectedPath"] = expectedPath });

    private static bool HasEntry(IEnumerable<string> entryNames, string expected)
        => entryNames.Any(name => string.Equals(name, expected, StringComparison.OrdinalIgnoreCase));

    private static bool IsSectionPart(string entryName)
        => entryName.StartsWith("Contents/section", StringComparison.OrdinalIgnoreCase)
            && entryName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    private static bool IsXmlLikePart(string entryName)
        => entryName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || entryName.EndsWith(".hpf", StringComparison.OrdinalIgnoreCase)
            || entryName.EndsWith(".rdf", StringComparison.OrdinalIgnoreCase);

    private static bool ContainsAny(string value, params string[] needles)
        => needles.Any(needle => value.Contains(needle, StringComparison.OrdinalIgnoreCase));

    private static bool IsPackageBlocking(ValidationError error, bool strictOrphans)
    {
        if (error.ErrorType.Equals("package_version_missing", StringComparison.OrdinalIgnoreCase))
            return false;
        if (!strictOrphans && error.ErrorType.Contains("bindata_orphan", StringComparison.OrdinalIgnoreCase))
            return false;
        if (error.Severity != IssueSeverity.Error) return false;
        return ContainsAny(
            error.ErrorType,
            "zip",
            "opf",
            "package",
            "mimetype",
            "container",
            "xml",
            "bindata");
    }
}
