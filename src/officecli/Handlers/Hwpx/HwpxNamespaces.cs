// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;

namespace OfficeCli.Handlers;

/// <summary>HWPX (OWPML) XML namespace constants.</summary>
/// <remarks>
/// CRITICAL: Use XNamespace (static readonly), NOT string const.
/// "Hp" not "HP" — matches OWPML spec prefix.
/// OPF namespace uses trailing slash (matches Hancom output); no-slash variant normalized via LegacyToCanonical.
/// </remarks>
public static class HwpxNs
{
    // Body sections (.xml files under Contents/)
    public static readonly XNamespace Hs = "http://www.hancom.co.kr/hwpml/2011/section";
    public static readonly XNamespace Hp = "http://www.hancom.co.kr/hwpml/2011/paragraph";

    // Header (.xml file at Contents/header.xml)
    public static readonly XNamespace Hh = "http://www.hancom.co.kr/hwpml/2011/head";

    // Core types (child elements inside header structures: margin children, etc.)
    public static readonly XNamespace Hc = "http://www.hancom.co.kr/hwpml/2011/core";

    // OPF manifest (META-INF/container.xml or mimetype OPF)
    // Real Hancom files use trailing slash; some tooling omits it — support both via LegacyToCanonical
    public static readonly XNamespace Opf = "http://www.idpf.org/2007/opf/";
    public static readonly XNamespace Dc  = "http://purl.org/dc/elements/1.1/";

    // Namespace URIs that appear in HWPML 2016 docs — must be normalized to 2011 before parsing
    public static readonly Dictionary<string, string> LegacyToCanonical = new()
    {
        ["http://www.hancom.co.kr/hwpml/2016/section"]    = "http://www.hancom.co.kr/hwpml/2011/section",
        ["http://www.hancom.co.kr/hwpml/2016/paragraph"]  = "http://www.hancom.co.kr/hwpml/2011/paragraph",
        ["http://www.hancom.co.kr/hwpml/2016/head"]       = "http://www.hancom.co.kr/hwpml/2011/head",
        ["http://www.hancom.co.kr/hwpml/2016/core"]       = "http://www.hancom.co.kr/hwpml/2011/core",
        ["http://www.hancom.co.kr/hwpml/2016/app"]        = "http://www.hancom.co.kr/hwpml/2011/app",
        // OPF: Hancom uses trailing slash; normalize no-slash variant
        ["http://www.idpf.org/2007/opf\""]                = "http://www.idpf.org/2007/opf/\"",
    };
}
