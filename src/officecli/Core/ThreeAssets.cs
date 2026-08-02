// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for Three.js asset URLs in generated HTML.
/// CONSISTENCY(katex-mirror): same sourcing policy as <c>KatexAssets</c> /
/// <c>MermaidImageRenderer</c> — own mirror first, public CDN as fallback.
/// The mirror hosts <c>build/three.module.js</c> plus the full
/// <c>examples/jsm/</c> subtree under one immutable versioned prefix, so the
/// importmap's <c>three/addons/</c> prefix mapping keeps working for any
/// addon module.
///
/// Import maps have no per-entry fallback mechanism, so the CDN retry lives
/// at the dynamic <c>import()</c> call site instead: on mirror failure the
/// loader script re-imports from jsdelivr's <c>/+esm</c> endpoint, whose
/// rewritten absolute internal imports bypass the (mirror-pointing) import
/// map entirely and share one module graph, keeping THREE instances
/// consistent between the core module and addons.
/// </summary>
internal static class ThreeAssets
{
    public const string Version = "0.170.0";

    private const string MirrorBase = "https://d.officecli.ai/assets/three-" + Version;
    private const string CdnBase = "https://cdn.jsdelivr.net/npm/three@" + Version;

    /// <summary>Import map JSON — mirror-first bare-specifier resolution.</summary>
    public static string ImportMapJson =>
        "{\"imports\":{\"three\":\"" + MirrorBase + "/build/three.module.js\","
        + "\"three/addons/\":\"" + MirrorBase + "/examples/jsm/\"}}";

    /// <summary>CDN fallback for the core module (self-contained ESM build).</summary>
    public static string CdnEsmThreeUrl => CdnBase + "/+esm";

    /// <summary>CDN fallback for GLTFLoader (internal imports rewritten to absolute CDN URLs).</summary>
    public static string CdnEsmGltfLoaderUrl => CdnBase + "/examples/jsm/loaders/GLTFLoader.js/+esm";
}
