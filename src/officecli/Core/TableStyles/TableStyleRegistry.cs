// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Style IDs are PowerPoint's stable public identifiers (factual data).

namespace OfficeCli.Core.TableStyles;

/// <summary>
/// Maps each of PowerPoint's 74 built-in table-style GUIDs to its family
/// template ("Medium-Style-2", "Light-Style-1", ...) and accent variant
/// ("Accent1".."Accent6" or "" for the neutral / dk1 variant).
///
/// Also exposes the historical short-name aliases used by users on the
/// command line (e.g. style=medium2 → Accent1 of Medium-Style-2). The short
/// names map only to the Accent1 variant of each family (historical
/// behaviour); other accents must be specified by GUID.
/// </summary>
public static class TableStyleRegistry
{
    /// <summary>
    /// GUID → (FamilyName, AccentName) for all 74 built-in styles.
    /// FamilyName is one of: Themed-Style-1/2, Light-Style-1/2/3,
    /// Medium-Style-1/2/3/4, Dark-Style-1/2. AccentName is "" or
    /// Accent1..Accent6 (Dark-Style-2 is sparse — only 1/3/5).
    /// </summary>
    public static readonly IReadOnlyDictionary<string, (string Family, string Accent)>
        ByGuid = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase)
    {
        // Themed-Style-1 (no-style; minimal frame + body fill)
        ["{2D5ABB26-0587-4C30-8999-92F81FD0307C}"] = ("Themed-Style-1", ""),
        ["{3C2FFA5D-87B4-456A-9821-1D502468CF0F}"] = ("Themed-Style-1", "Accent1"),
        ["{284E427A-3D55-4303-BF80-6455036E1DE7}"] = ("Themed-Style-1", "Accent2"),
        ["{69C7853C-536D-4A76-A0AE-DD22124D55A5}"] = ("Themed-Style-1", "Accent3"),
        ["{775DCB02-9BB8-47FD-8907-85C794F793BA}"] = ("Themed-Style-1", "Accent4"),
        ["{35758FB7-9AC5-4552-8A53-C91805E547FA}"] = ("Themed-Style-1", "Accent5"),
        ["{08FB837D-C827-4EFA-A057-4D05807E0F7C}"] = ("Themed-Style-1", "Accent6"),

        // Themed-Style-2 (themed border on header + outer)
        ["{5940675A-B579-460E-94D1-54222C63F5DA}"] = ("Themed-Style-2", ""),
        ["{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}"] = ("Themed-Style-2", "Accent1"),
        ["{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}"] = ("Themed-Style-2", "Accent2"),
        ["{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}"] = ("Themed-Style-2", "Accent3"),
        ["{E269D01E-BC32-4049-B463-5C60D7B0CCD2}"] = ("Themed-Style-2", "Accent4"),
        ["{327F97BB-C833-4FB7-BDE5-3F7075034690}"] = ("Themed-Style-2", "Accent5"),
        ["{638B1855-1B75-4FBE-930C-398BA8C253C6}"] = ("Themed-Style-2", "Accent6"),

        // Light-Style-1 (clean grid, header underline, banded rows)
        ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = ("Light-Style-1", ""),
        ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = ("Light-Style-1", "Accent1"),
        ["{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}"] = ("Light-Style-1", "Accent2"),
        ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = ("Light-Style-1", "Accent3"),
        ["{D27102A9-8310-4765-A935-A1911B00CA55}"] = ("Light-Style-1", "Accent4"),
        ["{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}"] = ("Light-Style-1", "Accent5"),
        ["{68D230F3-CF80-4859-8CE7-A43EE81993B5}"] = ("Light-Style-1", "Accent6"),

        // Light-Style-2 (accent-tinted grid + light banded)
        ["{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}"] = ("Light-Style-2", ""),
        ["{69012ECD-51FC-41F1-AA8D-1B2483CD663E}"] = ("Light-Style-2", "Accent1"),
        ["{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}"] = ("Light-Style-2", "Accent2"),
        ["{F2DE63D5-997A-4646-A377-4702673A728D}"] = ("Light-Style-2", "Accent3"),
        ["{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}"] = ("Light-Style-2", "Accent4"),
        ["{5A111915-BE36-4E01-A7E5-04B1672EAD32}"] = ("Light-Style-2", "Accent5"),
        ["{912C8C85-51F0-491E-9774-3900AFEF0FD7}"] = ("Light-Style-2", "Accent6"),

        // Light-Style-3 (no internal grid, only header + outer)
        ["{616DA210-FB5B-4158-B5E0-FEB733F419BA}"] = ("Light-Style-3", ""),
        ["{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}"] = ("Light-Style-3", "Accent1"),
        ["{5DA37D80-6434-44D0-A028-1B22A696006F}"] = ("Light-Style-3", "Accent2"),
        ["{8799B23B-EC83-4686-B30A-512413B5E67A}"] = ("Light-Style-3", "Accent3"),
        ["{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}"] = ("Light-Style-3", "Accent4"),
        ["{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}"] = ("Light-Style-3", "Accent5"),
        ["{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}"] = ("Light-Style-3", "Accent6"),

        // Medium-Style-1
        ["{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}"] = ("Medium-Style-1", ""),
        ["{B301B821-A1FF-4177-AEE7-76D212191A09}"] = ("Medium-Style-1", "Accent1"),
        ["{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}"] = ("Medium-Style-1", "Accent2"),
        ["{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}"] = ("Medium-Style-1", "Accent3"),
        ["{1E171933-4619-4E11-9A3F-F7608DF75F80}"] = ("Medium-Style-1", "Accent4"),
        ["{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}"] = ("Medium-Style-1", "Accent5"),
        ["{10A1B5D5-9B99-4C35-A422-299274C87663}"] = ("Medium-Style-1", "Accent6"),

        // Medium-Style-2 (white grout banded tile look — most common style)
        ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = ("Medium-Style-2", ""),
        ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = ("Medium-Style-2", "Accent1"),
        ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"] = ("Medium-Style-2", "Accent2"),
        ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = ("Medium-Style-2", "Accent3"),
        ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = ("Medium-Style-2", "Accent4"),
        ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}"] = ("Medium-Style-2", "Accent5"),
        ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = ("Medium-Style-2", "Accent6"),

        // Medium-Style-3 (header underline + last-row underline only)
        ["{8EC20E35-A176-4012-BC5E-935CFFF8708E}"] = ("Medium-Style-3", ""),
        ["{6E25E649-3F16-4E02-A733-19D2CDBF48F0}"] = ("Medium-Style-3", "Accent1"),
        ["{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}"] = ("Medium-Style-3", "Accent2"),
        ["{EB344D84-9AFB-497E-A393-DC336BA19D2E}"] = ("Medium-Style-3", "Accent3"),
        ["{EB9631B5-78F2-41C9-869B-9F39066F8104}"] = ("Medium-Style-3", "Accent4"),
        ["{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}"] = ("Medium-Style-3", "Accent5"),
        ["{2A488322-F2BA-4B5B-9748-0D474271808F}"] = ("Medium-Style-3", "Accent6"),

        // Medium-Style-4 (outer box only, no internal lines)
        ["{D7AC3CCA-C797-4891-BE02-D94E43425B78}"] = ("Medium-Style-4", ""),
        ["{69CF1AB2-1976-4502-BF36-3FF5EA218861}"] = ("Medium-Style-4", "Accent1"),
        ["{8A107856-5554-42FB-B03E-39F5DBC370BA}"] = ("Medium-Style-4", "Accent2"),
        ["{0505E3EF-67EA-436B-97B2-0124C06EBD24}"] = ("Medium-Style-4", "Accent3"),
        ["{C4B1156A-380E-4F78-BDF5-A606A8083BF9}"] = ("Medium-Style-4", "Accent4"),
        ["{22838BEF-8BB2-4498-84A7-C5851F593DF1}"] = ("Medium-Style-4", "Accent5"),
        ["{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}"] = ("Medium-Style-4", "Accent6"),

        // Dark-Style-1 (dark tinted bands, no internal grid)
        ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = ("Dark-Style-1", ""),
        ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = ("Dark-Style-1", "Accent1"),
        ["{37CE84F3-28C3-443E-9E96-99CF82512B78}"] = ("Dark-Style-1", "Accent2"),
        ["{D03447BB-5D67-496B-8E87-E561075AD55C}"] = ("Dark-Style-1", "Accent3"),
        ["{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"] = ("Dark-Style-1", "Accent4"),
        ["{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}"] = ("Dark-Style-1", "Accent5"),
        ["{AF606853-7671-496A-8E4F-DF71F8EC918B}"] = ("Dark-Style-1", "Accent6"),

        // Dark-Style-2 (sparse — only neutral + Accent1/3/5)
        ["{5202B0CA-FC54-4496-8BCA-5EF66A818D29}"] = ("Dark-Style-2", ""),
        ["{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}"] = ("Dark-Style-2", "Accent1"),
        ["{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}"] = ("Dark-Style-2", "Accent3"),
        ["{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}"] = ("Dark-Style-2", "Accent5"),
    };

    /// <summary>
    /// Reverse map: (Family, Accent) → GUID. Built lazily from ByGuid; used
    /// when users write `style=medium2` on the CLI and we need to round-trip
    /// to the canonical GUID.
    /// </summary>
    public static readonly IReadOnlyDictionary<(string Family, string Accent), string>
        ByFamilyAccent = ByGuid
            .GroupBy(kv => kv.Value)
            .ToDictionary(g => g.Key, g => g.First().Key);

    /// <summary>
    /// CLI short names ("medium2", "light1", ...) → canonical GUID.
    /// Each short name maps to the *no-accent* (neutral / dk1) variant of
    /// the matching family. To pick an accent variant, use the compound
    /// form "<family>-accent<N>" (e.g. "dark2-accent1", "medium3-accent4");
    /// see <see cref="Resolve"/> for the parser. Keep this table aligned
    /// with <c>_tableStyleNameToGuid</c> in
    /// <c>PowerPointHandler.Helpers.cs</c> (the input alias map).
    /// Round-tripping via GuidToShortName preserves the CLI name.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, string> ByShortName =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["medium1"] = "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}",
        ["medium2"] = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
        ["medium3"] = "{8EC20E35-A176-4012-BC5E-935CFFF8708E}",
        ["medium4"] = "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
        ["light1"]  = "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
        ["light2"]  = "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}",
        ["light3"]  = "{616DA210-FB5B-4158-B5E0-FEB733F419BA}",
        ["dark1"]   = "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
        ["dark2"]   = "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}",
        ["none"]    = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
    };

    /// <summary>
    /// Resolve any input (GUID, short name "medium2", compound "dark2-accent1",
    /// or family name "Dark-Style-2") to its (family, accent) pair. Returns
    /// null if unknown.
    /// </summary>
    public static (string Family, string Accent)? Resolve(string? styleIdOrName)
    {
        if (string.IsNullOrWhiteSpace(styleIdOrName)) return null;
        var key = styleIdOrName.Trim();
        if (ByGuid.TryGetValue(key, out var pair)) return pair;
        if (ByShortName.TryGetValue(key, out var guid) && ByGuid.TryGetValue(guid, out pair))
            return pair;
        if (TryParseCompound(key, out var family, out var accent)
            && ByFamilyAccent.TryGetValue((family, accent), out _))
            return (family, accent);
        return null;
    }

    /// <summary>
    /// Resolve short name, GUID, or compound "<short>-accent<N>" to canonical
    /// GUID. Returns null when nothing matches.
    /// </summary>
    public static string? ShortNameToGuid(string? shortName)
    {
        if (string.IsNullOrWhiteSpace(shortName)) return null;
        var key = shortName.Trim();
        if (ByShortName.TryGetValue(key, out var g)) return g;
        if (TryParseCompound(key, out var family, out var accent)
            && ByFamilyAccent.TryGetValue((family, accent), out var guid))
            return guid;
        return null;
    }

    /// <summary>
    /// Reverse of <see cref="ByShortName"/>: GUID → CLI short name (e.g.
    /// "dark2" for the no-accent variant, "dark2-accent1" for the Accent1
    /// variant). Returns null for GUIDs outside the 74-entry catalogue.
    /// </summary>
    public static string? GuidToShortName(string guid)
    {
        if (string.IsNullOrEmpty(guid)) return null;
        foreach (var kv in ByShortName)
            if (string.Equals(kv.Value, guid, StringComparison.OrdinalIgnoreCase))
                return kv.Key;
        if (!ByGuid.TryGetValue(guid, out var pair)) return null;
        if (string.IsNullOrEmpty(pair.Accent)) return null;
        foreach (var kv in ByShortName)
        {
            if (!ByGuid.TryGetValue(kv.Value, out var basePair)) continue;
            if (basePair.Family == pair.Family && string.IsNullOrEmpty(basePair.Accent))
                return $"{kv.Key}-{pair.Accent.ToLowerInvariant()}";
        }
        return null;
    }

    private static bool TryParseCompound(string input, out string family, out string accent)
    {
        family = ""; accent = "";
        var dash = input.LastIndexOf('-');
        if (dash <= 0 || dash >= input.Length - 1) return false;
        var head = input[..dash];
        var tail = input[(dash + 1)..];
        if (!tail.StartsWith("accent", StringComparison.OrdinalIgnoreCase)) return false;
        if (!int.TryParse(tail.AsSpan(6), out var n) || n < 1 || n > 6) return false;
        if (!ByShortName.TryGetValue(head, out var baseGuid)) return false;
        if (!ByGuid.TryGetValue(baseGuid, out var basePair)) return false;
        family = basePair.Family;
        accent = "Accent" + n;
        return true;
    }
}
