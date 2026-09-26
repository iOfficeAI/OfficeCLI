// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Concurrent;
using System.Reflection;
using DocumentFormat.OpenXml;

namespace OfficeCli.Core;

/// <summary>
/// SDK-authoritative schema-order placement for children added via the generic
/// reflection fallbacks (<see cref="GenericXmlQuery.TryCreateTypedChild"/>,
/// <see cref="TypedAttributeFallback"/>) and the hand-rolled run/paragraph
/// property inserters.
///
/// <para>
/// The OpenXML SDK's <c>AppendChild</c> only appends, and even
/// <c>OpenXmlCompositeElement.AddChild</c> does not reliably reposition. Strict
/// consumers (and OOXML schema validation) reject children that sit out of the
/// CT_* sequence — e.g. <c>&lt;w:kern&gt;</c> after <c>&lt;w:sz&gt;</c> in rPr,
/// or <c>&lt;w:pStyle&gt;</c> not first in pPr. PowerPoint silently drops
/// out-of-order DrawingML elements.
/// </para>
///
/// <para>
/// Rather than hand-maintain per-container order arrays (CT_RPr alone has ~40
/// children; the prior <c>InsertRunPropInSchemaOrder</c> switch silently
/// dropped <c>kern</c>/<c>w</c>/<c>position</c> to "append at end"), this reads
/// the SDK's own compiled particle comparator (<c>CompiledParticle.Compare</c>)
/// via reflection — the exact ordering the SDK validator uses. It is complete
/// by construction and stays correct as the schema evolves. See
/// <c>DrawingEffectsHelper.InsertEffectInSchemaOrder</c> for the hand-array
/// precedent this generalizes.
/// </para>
///
/// <para>
/// Placement is MINIMAL and SAFE: only the newly added child is moved, to the
/// first slot where it precedes an existing sibling. Existing children —
/// including foreign / unknown elements (<c>mc:AlternateContent</c>, extension
/// lists) that the comparator can't order — are never reordered. The comparator
/// sorts unknown elements to the front, so a blind full sort would corrupt
/// extension placement; single-child placement avoids that entirely.
/// </para>
/// </summary>
internal static class SchemaOrder
{
    // Per-parent-type comparator cache. The compiled particle is static per
    // element type and the comparison depends only on the child element types,
    // so caching by the parent's runtime type is safe across instances and
    // threads. A null entry means the type has no usable particle (leaf
    // elements, reflection shape changed) — callers then leave order as-is.
    private static readonly ConcurrentDictionary<Type, Comparison<OpenXmlElement>?> s_cmpCache = new();

    // Per-(parent type, child qualified-name) schema maxOccurs cache. Keyed by
    // parent runtime type + "{ns}:{localName}" of the child. Value: the child's
    // maxOccurs in the parent's content model — 1 = singleton, anything else
    // (0 = unbounded, or a finite bound > 1) = repeatable, int.MinValue = the
    // child/particle couldn't be resolved (caller treats as repeatable).
    private static readonly ConcurrentDictionary<(Type, string), int> s_maxOccursCache = new();

    public const int MaxOccursUnknown = int.MinValue;

    /// <summary>
    /// The schema maxOccurs of <paramref name="child"/> within
    /// <paramref name="parent"/>'s content model, via the SDK's compiled
    /// particle. Returns 1 for a singleton child, 0 for unbounded, a finite
    /// bound for bounded-repeatable, or <see cref="MaxOccursUnknown"/> when the
    /// particle can't be read or the child isn't found (caller should then NOT
    /// assume singleton). Used by the generic add to decide replace (singleton)
    /// vs append-distinct-sibling (repeatable).
    /// </summary>
    [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2075",
        Justification = "Reflecting the SDK's internal element Metadata.Particle, same members the validator uses; returns MaxOccursUnknown if trimmed.")]
    public static int GetMaxOccurs(OpenXmlElement parent, OpenXmlElement child)
    {
        var key = (parent.GetType(), $"{child.NamespaceUri}:{child.LocalName}");
        return s_maxOccursCache.GetOrAdd(key, static (k, p) =>
        {
            try
            {
                var metaProp = typeof(OpenXmlElement).GetProperty("Metadata",
                    BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                var meta = metaProp?.GetValue(p);
                var compiled = meta?.GetType().GetProperty("Particle")?.GetValue(meta);
                var root = compiled?.GetType().GetProperty("Particle")?.GetValue(compiled);
                if (root == null) return MaxOccursUnknown;
                return WalkForMaxOccurs(root, k.Item2);
            }
            catch { return MaxOccursUnknown; }
        }, parent);
    }

    [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2075",
        Justification = "Reflecting the SDK's internal compiled-particle members (ParticleType/Type/MaxOccurs/ChildrenParticles), same members the validator uses; returns MaxOccursUnknown if trimmed. Sibling of GetMaxOccurs/GetComparison, which carry the same suppression.")]
    private static int WalkForMaxOccurs(object pc, string childQName)
    {
        var t = pc.GetType();
        var particleType = t.GetProperty("ParticleType")?.GetValue(pc)?.ToString();
        if (particleType == "Element")
        {
            // OpenXmlSchemaType.ToString() == "{ct-ns}:CT_Name/{el-ns}:localName".
            // Both ns segments are full URLs (containing '/'), so the element id
            // "{el-ns}:localName" == childQName is a SUFFIX of the string — match
            // by EndsWith, not by splitting on '/'.
            var typeStr = t.GetProperty("Type")?.GetValue(pc)?.ToString() ?? "";
            if (typeStr.EndsWith(childQName, StringComparison.Ordinal))
                return (int)(t.GetProperty("MaxOccurs")?.GetValue(pc) ?? MaxOccursUnknown);
            return MaxOccursUnknown;
        }
        if (t.GetProperty("ChildrenParticles")?.GetValue(pc) is System.Collections.IEnumerable kids)
        {
            foreach (var k in kids)
            {
                if (k == null) continue;
                var r = WalkForMaxOccurs(k, childQName);
                if (r != MaxOccursUnknown) return r;
            }
        }
        return MaxOccursUnknown;
    }

    /// <summary>
    /// Move <paramref name="child"/> — already a child of <paramref name="parent"/>,
    /// typically just appended — to its schema-correct slot: immediately before
    /// the first existing sibling it should precede per the SDK particle order.
    /// No-op when the parent is not composite, has no compiled particle, the
    /// comparator can't order the child, or nothing should follow it.
    /// </summary>
    public static void Place(OpenXmlElement parent, OpenXmlElement child)
    {
        if (parent is not OpenXmlCompositeElement) return;
        if (!ReferenceEquals(child.Parent, parent)) return;

        var cmp = GetComparison(parent);
        if (cmp == null) return;

        OpenXmlElement? anchor = null;
        foreach (var sibling in parent.ChildElements)
        {
            if (ReferenceEquals(sibling, child)) continue;
            int order;
            try { order = cmp(child, sibling); }
            catch { return; } // comparator failure → leave as-is (already appended)
            if (order < 0) { anchor = sibling; break; }
        }

        // anchor == null: child belongs at/after the current tail → leave it
        // where it was appended. Otherwise hoist it before the first sibling it
        // precedes. The SDK's InsertBefore refuses an element that is still
        // attached ("part of a tree"), so detach first, then re-insert.
        if (anchor != null && !ReferenceEquals(child.NextSibling(), anchor))
        {
            child.Remove();
            parent.InsertBefore(child, anchor);
        }
    }

    [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2075",
        Justification = "Reflecting over the SDK's internal element Metadata.Particle to reach the public CompiledParticle.Compare comparator. These members back the SDK's own validation and are preserved; if trimmed away, the probe returns null and callers leave child order unchanged (the prior, also-valid-enough behavior).")]
    private static Comparison<OpenXmlElement>? GetComparison(OpenXmlElement parent)
    {
        return s_cmpCache.GetOrAdd(parent.GetType(), static (type, p) =>
        {
            try
            {
                // Metadata is an instance property on OpenXmlElement; the
                // particle it exposes is identical across instances of the same
                // type, so building the delegate from any live instance and
                // caching by type is correct. For the containers listed in
                // s_particleAuthority the delegate is built from a different
                // type's particle, deliberately — see that field.
                var source = s_particleAuthority.TryGetValue(type, out var authority)
                    ? authority()
                    : p;
                var metaProp = typeof(OpenXmlElement).GetProperty("Metadata",
                    BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                var meta = metaProp?.GetValue(source);
                var particle = meta?.GetType().GetProperty("Particle")?.GetValue(meta);
                if (particle == null) return null;
                var compare = particle.GetType().GetMethod("Compare",
                    new[] { typeof(OpenXmlElement), typeof(OpenXmlElement) });
                if (compare == null) return null;
                return (a, b) => (int)compare.Invoke(particle, new object[] { a, b })!;
            }
            catch
            {
                return null;
            }
        }, parent);
    }

    // Per-parent-type particle AUTHORITY override — for the containers whose own
    // compiled particle does not model the schema type the validator tests them
    // against. w:style/w:rPr is the case that bites: ISO/IEC 29500-4 CT_Style
    // declares <xsd:element name="rPr" type="CT_RPr" minOccurs="0"/> — the same
    // complex type a run's w:rPr uses — while SDK 3.4.1's compiled particle for
    // StyleRunProperties is a TRUNCATED CT_RPr: 35 children against
    // RunProperties' 51, missing rStyle, highlight, rtl, cs, oMath and the w14
    // extension block (its typed class has no Highlight/RunStyle/... property
    // either). The comparator sorts what it cannot recognise to the FRONT, so
    // placing <w:highlight/> in a style's rPr hoisted it ahead of <w:rFonts/> —
    // a slot neither CT_RPr nor Word's own writer produces. Borrow the
    // RunProperties particle for that container: for every element both
    // particles declare the relative order is identical, so this only moves the
    // five undeclared ones and the w14 block, and it leaves GetMaxOccurs (which
    // reads the parent's own particle) untouched.
    private static readonly ConcurrentDictionary<Type, Func<OpenXmlElement>> s_particleAuthority = new()
    {
        [typeof(DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties)] =
            static () => new DocumentFormat.OpenXml.Wordprocessing.RunProperties(),
    };
}
