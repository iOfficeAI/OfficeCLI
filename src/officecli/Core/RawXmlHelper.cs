// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
// OOXML0001: PackageExtensions.GetPackage() is marked [Experimental] but it
// is the only public path to the underlying IPackage. The previous reflection
// workaround tripped trim analysis (IL2026/IL2075/IL2060) and was fragile —
// IPackage/IPackagePart are themselves public interfaces with public methods,
// so once GetPackage() returns we are entirely on the public surface.
#pragma warning disable OOXML0001
using DocumentFormat.OpenXml.Experimental;
#pragma warning restore OOXML0001

namespace OfficeCli.Core;

/// <summary>
/// Shared helper for raw XML operations (read/write via XPath).
/// This enables AI to perform any OpenXML operation by manipulating XML directly.
/// </summary>
internal static class RawXmlHelper
{
    /// <summary>
    /// Perform a raw XML operation on a document part's root element.
    /// </summary>
    /// <param name="rootElement">The OpenXml root element (e.g. Document, Worksheet, Slide)</param>
    /// <param name="xpath">XPath expression to locate target element(s)</param>
    /// <param name="action">Operation: append, prepend, insertbefore, insertafter, replace, remove, setattr</param>
    /// <param name="xml">XML fragment for append/prepend/insert/replace, or attr=value for setattr</param>
    /// <returns>Number of elements affected</returns>
    public static int Execute(OpenXmlPartRootElement rootElement, string xpath, string action, string? xml)
    {
        var (xDoc, affected) = ExecuteOnXmlString(rootElement.OuterXml, xpath, action, xml);
        if (affected == 0) return 0;
        // Write modified XML back to the OpenXml element.
        // Propagate namespace declarations from root to direct child elements,
        // so that each child's ToString() produces self-contained XML with
        // all necessary namespace bindings (otherwise inherited namespaces are lost).
        var rootNsAttrs = xDoc.Root!.Attributes()
            .Where(a => a.IsNamespaceDeclaration).ToList();
        foreach (var child in xDoc.Root.Elements())
        {
            foreach (var nsAttr in rootNsAttrs)
            {
                if (child.Attribute(nsAttr.Name) == null)
                    child.SetAttributeValue(nsAttr.Name, nsAttr.Value);
            }
        }
        rootElement.InnerXml = string.Concat(xDoc.Root.Nodes().Select(n => n.ToString()));
        // The InnerXml setter restores inner content but does NOT touch root
        // attributes — so non-xmlns attrs like `mc:Ignorable` carried by the
        // replacement root would be silently lost on round-trip. Copy them
        // through. Skip xmlns declarations (SDK manages those via its typed
        // prefix table). Need the prefix for namespaced attrs so SDK renders
        // them as `<prefix>:<localname>` rather than re-prefixing under the
        // SDK's auto-generated alias.
        foreach (var attr in xDoc.Root.Attributes())
        {
            if (attr.IsNamespaceDeclaration) continue;
            var ns = attr.Name.NamespaceName;
            // Resolve the prefix as the source XML had it (walk parent's
            // in-scope namespace declarations). XLinq's XAttribute.Name has
            // no Prefix property, only the parent element does.
            var prefix = string.IsNullOrEmpty(ns)
                ? string.Empty
                : (xDoc.Root.GetPrefixOfNamespace(ns) ?? string.Empty);
            var openXmlAttr = new DocumentFormat.OpenXml.OpenXmlAttribute(
                prefix, attr.Name.LocalName, ns, attr.Value);
            rootElement.SetAttribute(openXmlAttr);
        }
        // For each Markup Compatibility extension prefix listed in the
        // post-mutation `mc:Ignorable` on the root, propagate the matching
        // xmlns declaration to the SDK root so the prefix-to-URI binding is
        // visible on the element that carries Ignorable. Without this, strict
        // consumers (PowerPoint) reject the file even though the SDK serializer
        // appears to round-trip the inner content — `SetAttribute` skipped
        // xmlns nodes above and the SDK won't synthesize a declaration for a
        // prefix it has no typed knowledge of.
        var ignorableValue = (string?)xDoc.Root.Attribute(McNs + "Ignorable");
        if (!string.IsNullOrWhiteSpace(ignorableValue))
        {
            foreach (var prefix in ignorableValue.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            {
                var ns = xDoc.Root.GetNamespaceOfPrefix(prefix);
                if (ns == null) continue;
                // Skip if SDK already has it declared (avoid duplicate decl).
                var existing = rootElement.LookupNamespace(prefix);
                if (existing == ns.NamespaceName) continue;
                try { rootElement.AddNamespaceDeclaration(prefix, ns.NamespaceName); }
                catch (InvalidOperationException) { /* already present under another binding */ }
            }
        }
        return affected;
    }

    /// <summary>
    /// Apply a raw XML operation directly on a part's stream (no SDK typed
    /// root needed). Used for arbitrary XML parts addressed by zip URI —
    /// sheet1.xml, footnotes.xml, customXml/item1.xml, etc.
    /// </summary>
    public static int Execute(OpenXmlPart part, string xpath, string action, string? xml)
    {
        var sourceXml = ReadPartXml(part);
        var (xDoc, affected) = ExecuteOnXmlString(sourceXml, xpath, action, xml);
        if (affected == 0) return 0;
        WritePartXml(part, xDoc.ToString(SaveOptions.DisableFormatting));
        return affected;
    }

    private static (XDocument xDoc, int affected) ExecuteOnXmlString(
        string sourceXml, string xpath, string action, string? xml)
    {
        var xDoc = XDocument.Parse(sourceXml);
        var nsManager = BuildNamespaceManager(xDoc);

        // Validate action before evaluating xpath so an unknown action is
        // never masked by a no-match xpath. Without this, a typo'd action
        // combined with a stale xpath would batch-report "OK" while doing
        // nothing — the user sees no signal that the action name was wrong.
        var normalizedAction = action.ToLowerInvariant();
        switch (normalizedAction)
        {
            case "append":
            case "prepend":
            case "insertbefore":
            case "before":
            case "insertafter":
            case "after":
            case "replace":
            case "remove":
            case "delete":
            case "setattr":
                break;
            default:
                throw new ArgumentException($"Unknown action: {action}. Supported: append, prepend, insertbefore, insertafter, replace, remove, setattr");
        }

        var nodes = xDoc.XPathSelectElements(xpath, nsManager).ToList();
        if (nodes.Count == 0)
        {
            // Throw rather than return 0 affected with a stderr nudge: the
            // stderr line is trivially dropped by pipelines and batch
            // envelopes, leaving callers with success:true on a no-op. A
            // typo'd xpath then looks indistinguishable from a real
            // mutation. Surface the failure where every consumer
            // (standalone, batch, resident) already handles exceptions.
            throw new ArgumentException(
                $"raw-set: XPath matched no elements: {xpath}. " +
                "Hint: auto-registered namespace prefixes: " +
                string.Join(", ", CommonNamespaces.Keys.Order()) +
                ". No xmlns declarations needed in --xml fragments.");
        }

        int affected = 0;

        foreach (var node in nodes)
        {
            switch (action.ToLowerInvariant())
            {
                case "append":
                    if (xml == null) throw new ArgumentException("--xml is required for append");
                    var appendFragment = ParseFragment(xml, xDoc);
                    foreach (var el in appendFragment)
                        node.Add(el);
                    affected++;
                    break;

                case "prepend":
                    if (xml == null) throw new ArgumentException("--xml is required for prepend");
                    var prependFragment = ParseFragment(xml, xDoc);
                    foreach (var el in prependFragment.AsEnumerable().Reverse())
                        node.AddFirst(el);
                    affected++;
                    break;

                case "insertbefore" or "before":
                    if (xml == null) throw new ArgumentException("--xml is required for insertbefore");
                    RequireParent(node, "insertbefore");
                    var beforeFragment = ParseFragment(xml, xDoc);
                    // AddBeforeSelf lands each element immediately before the
                    // anchor, i.e. AFTER everything inserted so far — forward
                    // iteration preserves source order. (The reverse idiom is
                    // insertafter-only; reversing here flipped a multi-element
                    // fragment, splitting bookmarkStart/End pairs so the id
                    // balancer synthesized a duplicate w:id end marker.)
                    foreach (var el in beforeFragment)
                        node.AddBeforeSelf(el);
                    affected++;
                    break;

                case "insertafter" or "after":
                    if (xml == null) throw new ArgumentException("--xml is required for insertafter");
                    RequireParent(node, "insertafter");
                    var afterFragment = ParseFragment(xml, xDoc);
                    // AddAfterSelf inserts immediately after `node`, so calling it
                    // repeatedly against the SAME anchor REVERSES a multi-element
                    // fragment (start,end → node,end,start). Iterate in REVERSE and
                    // keep anchoring to `node`, mirroring insertbefore — each element
                    // lands right after `node`, yielding source order. (Chaining the
                    // anchor off the just-added node does NOT work: AddAfterSelf clones
                    // a parented element, so the loop variable still points at the
                    // detached fragment node, the chain breaks, and only the first
                    // element reaches the document — silently dropping the rest of a
                    // multi-marker fragment, e.g. a second tr-level bookmark.) A
                    // reversed start/end pair also desynced the id-balancer into
                    // duplicate bookmark ids.
                    foreach (var el in afterFragment.AsEnumerable().Reverse())
                        node.AddAfterSelf(el);
                    affected++;
                    break;

                case "replace":
                    if (xml == null) throw new ArgumentException("--xml is required for replace");
                    var replaceFragment = ParseFragment(xml, xDoc);
                    node.ReplaceWith(replaceFragment.ToArray());
                    affected++;
                    break;

                case "remove" or "delete":
                    RequireParent(node, "remove");
                    node.Remove();
                    affected++;
                    break;

                case "setattr":
                    if (xml == null) throw new ArgumentException("--xml is required for setattr (format: name=value)");
                    var eqIdx = xml.IndexOf('=');
                    if (eqIdx <= 0) throw new ArgumentException("setattr format: name=value");
                    var attrName = xml[..eqIdx];
                    var attrValue = xml[(eqIdx + 1)..];

                    // Handle namespaced attributes (e.g. w:val)
                    var colonIdx = attrName.IndexOf(':');
                    if (colonIdx > 0)
                    {
                        var prefix = attrName[..colonIdx];
                        var localName = attrName[(colonIdx + 1)..];
                        var ns = nsManager.LookupNamespace(prefix);
                        if (ns != null)
                            node.SetAttributeValue(XName.Get(localName, ns), attrValue);
                        else
                            node.SetAttributeValue(attrName, attrValue);
                    }
                    else
                    {
                        node.SetAttributeValue(attrName, attrValue);
                    }
                    affected++;
                    break;

                default:
                    throw new ArgumentException($"Unknown action: {action}. Supported: append, prepend, insertbefore, insertafter, replace, remove, setattr");
            }
        }

        // mc:AlternateContent / mc:Choice fragments injected via raw-set carry
        // their own inline `xmlns:mc` and `xmlns:<ext>` declarations, but the
        // Markup Compatibility spec (ECMA-376 Part 3) requires every extension
        // prefix referenced by `mc:Choice/@Requires` to also be listed in an
        // ancestor `mc:Ignorable` attribute. Without it, strict consumers
        // (PowerPoint, Word strict mode) reject the file at load time even
        // though the SDK validator passes. PowerPoint's own files always
        // declare e.g. `mc:Ignorable="p159"` on <p:sld>; replay-from-dump must
        // mirror this. Sweep the post-mutation tree for any `mc:Choice@Requires`
        // and merge the prefixes into the root's mc:Ignorable (de-duplicated).
        EnsureMcIgnorableForChoiceRequires(xDoc);

        // Schema-order known well-known containers whose children land in a
        // fixed sequence. Raw-set callers (notably the pptx dump replay for
        // exotic transition/timing) emit an `append` on the parent without
        // knowing the replay-time sibling state — a later-emitted
        // `<p:transition>` would otherwise sit AFTER an earlier semantically
        // added `<p:timing>`, which strict consumers reject. The reorder is
        // a no-op when children are already in spec order, so existing
        // raw-set callers stay byte-stable.
        EnforceKnownSchemaOrder(xDoc);

        return (xDoc, affected);
    }

    // Map (parent-namespace, parent-localName) → ordered list of expected
    // child local names. Only well-known containers whose children are
    // affected by raw-set `append` callers; expand as new dump-replay paths
    // surface drift. Children whose localName is not in the list are kept in
    // their existing relative position at the tail.
    private static readonly Dictionary<(string ns, string name), string[]> SchemaOrderMap =
        new()
        {
            // p:sld — ECMA-376 CT_Slide
            [("http://schemas.openxmlformats.org/presentationml/2006/main", "sld")] =
                new[] { "cSld", "clrMapOvr", "transition", "timing", "extLst" },
            // p:sldLayout — CT_SlideLayout
            [("http://schemas.openxmlformats.org/presentationml/2006/main", "sldLayout")] =
                new[] { "cSld", "clrMapOvr", "transition", "timing", "hf", "extLst" },
            // p:sldMaster — CT_SlideMaster
            [("http://schemas.openxmlformats.org/presentationml/2006/main", "sldMaster")] =
                new[] { "cSld", "clrMap", "sldLayoutIdLst", "transition", "timing", "hf", "txStyles", "extLst" },
            // p:sp — CT_Shape: nvSpPr, spPr, style?, txBody?, extLst?
            // The dump->replay path raw-set appends <p:style> after the shape
            // is created with its txBody; without reorder, style lands AFTER
            // txBody and strict consumers reject the file.
            [("http://schemas.openxmlformats.org/presentationml/2006/main", "sp")] =
                new[] { "nvSpPr", "spPr", "style", "txBody", "extLst" },
        };

    private static void EnforceKnownSchemaOrder(XDocument xDoc)
    {
        if (xDoc.Root == null) return;
        // Apply to the document root and any nested matches (cheap walk).
        foreach (var el in xDoc.Descendants().Prepend(xDoc.Root))
        {
            if (!SchemaOrderMap.TryGetValue((el.Name.NamespaceName, el.Name.LocalName), out var order))
                continue;
            var children = el.Elements().ToList();
            // Order index for known names; unknown children sort last in
            // their current order (preserves source order for unknowns —
            // mc:AlternateContent wrappers carrying transition/timing live
            // under a different localName so they fall here, but they
            // were already placed by the caller).
            int IndexOf(XElement c)
            {
                // mc:AlternateContent that wraps a transition or timing
                // should sort to the slot of the wrapped element.
                if (c.Name.LocalName == "AlternateContent" &&
                    c.Name.NamespaceName == McNs.NamespaceName)
                {
                    foreach (var nested in c.Descendants())
                    {
                        var idx = Array.IndexOf(order, nested.Name.LocalName);
                        if (idx >= 0) return idx;
                    }
                }
                var i = Array.IndexOf(order, c.Name.LocalName);
                return i >= 0 ? i : order.Length;
            }
            // Stable sort by schema order, preserving source order within ties.
            var indexed = children.Select((c, i) => (child: c, slot: IndexOf(c), origIdx: i))
                .OrderBy(t => t.slot).ThenBy(t => t.origIdx).ToList();
            // Skip work if already in order.
            bool needsReorder = false;
            for (int i = 0; i < children.Count; i++)
            {
                if (!ReferenceEquals(children[i], indexed[i].child)) { needsReorder = true; break; }
            }
            if (!needsReorder) continue;
            foreach (var c in children) c.Remove();
            foreach (var t in indexed) el.Add(t.child);
        }
    }

    private static readonly XNamespace McNs =
        "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static void EnsureMcIgnorableForChoiceRequires(XDocument xDoc)
    {
        var root = xDoc.Root;
        if (root == null) return;

        // Collect every prefix listed in mc:Choice/@Requires (whitespace-separated).
        // Spec: Requires may name multiple prefixes; treat each token independently.
        var required = new HashSet<string>(StringComparer.Ordinal);
        foreach (var choice in root.Descendants(McNs + "Choice"))
        {
            var req = (string?)choice.Attribute("Requires");
            if (string.IsNullOrWhiteSpace(req)) continue;
            foreach (var token in req.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                required.Add(token);
        }
        if (required.Count == 0) return;

        // For each required prefix, the extension namespace declaration must be
        // visible in the document. Inline `xmlns:p159="..."` on the descendant
        // `<p:transition>` is in-scope for the Choice but is NOT what makes
        // mc:Ignorable lookup valid — strict consumers want the prefix-to-uri
        // binding resolvable from the element that *carries* mc:Ignorable.
        // Copy any inline declaration of that prefix up to the root if missing.
        foreach (var prefix in required)
        {
            if (root.GetNamespaceOfPrefix(prefix) != null) continue;
            // Walk descendants for the first inline declaration of this prefix.
            var declared = root.Descendants()
                .SelectMany(e => e.Attributes())
                .FirstOrDefault(a => a.IsNamespaceDeclaration && a.Name.LocalName == prefix);
            if (declared != null)
                root.SetAttributeValue(XNamespace.Xmlns + prefix, declared.Value);
        }

        // Ensure mc namespace itself is declared on the root so the Ignorable
        // attribute name binds correctly.
        if (root.GetPrefixOfNamespace(McNs) == null)
            root.SetAttributeValue(XNamespace.Xmlns + "mc", McNs.NamespaceName);

        // Merge required prefixes into mc:Ignorable, preserving any pre-existing
        // tokens and de-duplicating. Stable token order: existing first, new appended.
        var ignorableAttr = root.Attribute(McNs + "Ignorable");
        var existing = (ignorableAttr?.Value ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .ToList();
        var existingSet = new HashSet<string>(existing, StringComparer.Ordinal);
        var changed = false;
        foreach (var prefix in required)
        {
            if (existingSet.Add(prefix))
            {
                existing.Add(prefix);
                changed = true;
            }
        }
        if (changed || ignorableAttr == null)
            root.SetAttributeValue(McNs + "Ignorable", string.Join(" ", existing));
    }

    private static readonly Dictionary<string, string> CommonNamespaces = new()
    {
        ["w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ["a"] = "http://schemas.openxmlformats.org/drawingml/2006/main",
        ["p"] = "http://schemas.openxmlformats.org/presentationml/2006/main",
        ["x"] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ["wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ["c"] = "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ["xdr"] = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        ["wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ["wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ["v"] = "urn:schemas-microsoft-com:vml",
    };

    /// <summary>
    /// Guard actions that need a parent element. An XPath like <c>//*</c> also
    /// matches the document root, whose <see cref="XElement.Parent"/> is null;
    /// calling Remove()/AddBeforeSelf()/AddAfterSelf() on it threw a raw
    /// NullReferenceException. Surface a clear, actionable error instead.
    /// </summary>
    private static void RequireParent(XElement node, string action)
    {
        if (node.Parent == null)
            throw new ArgumentException(
                $"Cannot {action} the document root element <{node.Name.LocalName}>. " +
                $"Target a child element with a more specific --xpath (the root has no parent).");
    }

    private static List<XElement> ParseFragment(string xml, XDocument contextDoc)
    {
        // Collect namespace declarations from the context document
        var nsDict = new Dictionary<string, string>(CommonNamespaces);
        string? defaultNs = null;

        if (contextDoc.Root != null)
        {
            // Inherit the default namespace from the document root so that
            // unprefixed elements (e.g. <mergeCells>) are parsed into the
            // correct namespace (e.g. spreadsheetml) instead of empty namespace.
            var rootNsName = contextDoc.Root.Name.NamespaceName;
            if (!string.IsNullOrEmpty(rootNsName))
                defaultNs = rootNsName;

            foreach (var attr in contextDoc.Root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                var prefix = attr.Name.Namespace == XNamespace.Xmlns ? attr.Name.LocalName : "";
                if (!string.IsNullOrEmpty(prefix))
                    nsDict[prefix] = attr.Value;
            }
        }

        var prefixedNs = string.Join(" ", nsDict.Select(kv => $"xmlns:{kv.Key}=\"{kv.Value}\""));
        var defaultNsDecl = !string.IsNullOrEmpty(defaultNs) ? $"xmlns=\"{defaultNs}\"" : "";
        var wrappedXml = $"<_root {defaultNsDecl} {prefixedNs}>{xml}</_root>";
        // BUG-DUMP-R35-2: parse whitespace-PRESERVED, then normalize. The
        // default loader (LoadOptions.None) silently drops every
        // whitespace-only text node — including content whitespace in leaf
        // text elements (`<w:t> </w:t>` without xml:space="preserve"), so a
        // space-only run injected via raw-set vanished from the document
        // ("John Smith" → "JohnSmith"). NormalizeFragmentWhitespace keeps the
        // old behavior for inter-element formatting whitespace (pretty-printed
        // --xml input stays clean) but retains leaf-content whitespace and
        // stamps xml:space="preserve" on its parent so the downstream SDK
        // parse (OpenXmlElement.InnerXml, which also discards insignificant
        // whitespace) and every later reopen keep it too.
        var parsed = XDocument.Parse(wrappedXml, LoadOptions.PreserveWhitespace);
        NormalizeFragmentWhitespace(parsed.Root!);
        return parsed.Root!.Elements().ToList();
    }

    // BUG-DUMP-R35-2: post-pass over a whitespace-preserved fragment parse.
    // A whitespace-only text node is FORMATTING when its parent also has
    // element children (or is the synthetic _root wrapper) — remove it,
    // matching the previous LoadOptions.None behavior. It is CONTENT when its
    // parent is a leaf element (`<w:t> </w:t>`, `<a:t> </a:t>`) — keep it and
    // stamp xml:space="preserve" (XML-core attribute, namespace-independent)
    // so both XLinq re-serialization and the OpenXml SDK reader treat it as
    // significant.
    private static void NormalizeFragmentWhitespace(XElement root)
    {
        var wsTextNodes = root.DescendantNodes().OfType<XText>()
            .Where(t => t.Value.Length > 0 && string.IsNullOrWhiteSpace(t.Value))
            .ToList();
        foreach (var t in wsTextNodes)
        {
            var parent = t.Parent;
            if (parent == null) continue;
            // xml:space="preserve" already in scope → significant by decree.
            var inScope = parent.AncestorsAndSelf()
                .Select(a => (string?)a.Attribute(XNamespace.Xml + "space"))
                .FirstOrDefault(v => v != null);
            if (string.Equals(inScope, "preserve", StringComparison.Ordinal))
                continue;
            if (parent == root || parent.Elements().Any())
            {
                t.Remove(); // formatting whitespace between elements
            }
            else
            {
                parent.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            }
        }
    }

    private static XmlNamespaceManager BuildNamespaceManager(XDocument xDoc)
    {
        var nsManager = new XmlNamespaceManager(new NameTable());

        if (xDoc.Root != null)
        {
            foreach (var attr in xDoc.Root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                var prefix = attr.Name.LocalName;
                if (attr.Name.Namespace == XNamespace.Xmlns)
                {
                    nsManager.AddNamespace(prefix, attr.Value);
                }
                else if (attr.Name == "xmlns")
                {
                    // Default namespace — assign a usable prefix
                    nsManager.AddNamespace("default", attr.Value);
                }
            }
        }

        // Ensure common OpenXML namespaces are available
        TryAddNamespace(nsManager, "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        TryAddNamespace(nsManager, "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        TryAddNamespace(nsManager, "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        TryAddNamespace(nsManager, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        TryAddNamespace(nsManager, "x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        TryAddNamespace(nsManager, "wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        TryAddNamespace(nsManager, "mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        TryAddNamespace(nsManager, "c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        TryAddNamespace(nsManager, "xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");

        return nsManager;
    }

    /// <summary>
    /// Validate an OpenXmlPackage and return structured errors.
    /// </summary>
    public static List<ValidationError> ValidateDocument(OpenXmlPackage package, string? filePath = null)
    {
        // Validate against a throwaway clone, never the live package.
        // PreflightXmlParts opens each part's stream (GetStream); on the LIVE
        // package that desyncs the SDK's DOM tracking for parts whose root
        // element is dirty-but-unflushed (e.g. a StylesPart just created by a
        // style-setting edit). The part then serializes EMPTY on the next Save
        // ("Root element is missing"). Clone(stream) snapshots the current
        // in-memory state into an independent package, so reading its streams
        // cannot harm the document the caller will go on to save. (Parameterless
        // Clone() needs an in-memory package-factory feature that isn't
        // registered here; the stream overload is the one used elsewhere.)
        // Call the TYPED Clone overload: cloning through the base OpenXmlPackage
        // reference needs an IPackageFactoryFeature that isn't registered for
        // packages opened via SpreadsheetDocument/Word.../Presentation.Open, but
        // the typed Clone(Stream, bool) works (same call HtmlPreview uses).
        //
        // BUT Clone(stream) is itself a stream read of every LIVE part — so it
        // re-introduces the exact desync the clone was meant to avoid: cloning a
        // package with a dirty-but-unflushed StylesPart desyncs that live part,
        // which then serializes EMPTY on the caller's next Save (a `set` that
        // creates styles + an in-session `validate` over the resident pipe = a
        // 0-byte styles.xml on close). Flush each ALREADY-LOADED part's DOM back
        // to its stream first, so both Clone and PreflightXmlParts read in-sync
        // bytes and never desync. Only loaded parts are flushed: an unloaded part
        // cannot be dirty, and force-loading it would make the caller's Save
        // re-serialize an untouched part (byte churn) — validate must stay
        // read-only. There is no public IsRootElementLoaded in OpenXml 3.4, so
        // FlushLoadedPartRoots peeks the private loaded-root field reflectively
        // rather than touching part.RootElement (which would trigger a load).
        FlushLoadedPartRoots(package);
        using var cloneStream = new MemoryStream();
        using OpenXmlPackage clone = package switch
        {
            DocumentFormat.OpenXml.Packaging.SpreadsheetDocument s => s.Clone(cloneStream, false),
            DocumentFormat.OpenXml.Packaging.WordprocessingDocument w => w.Clone(cloneStream, false),
            DocumentFormat.OpenXml.Packaging.PresentationDocument p => p.Clone(cloneStream, false),
            _ => package.Clone(cloneStream, false),
        };
        var errors = PreflightXmlParts(clone);
        // BUG-R5B(BUG2): an orphaned header/footer reference (a
        // w:headerReference / w:footerReference whose r:id has no matching
        // relationship in its part) makes the SDK validator NRE before it can
        // produce any results — the user gets a bare "NullReferenceException"
        // instead of an actionable diagnostic. Detect these ourselves and emit a
        // clean schema-style error naming the dangling r:id, mirroring the
        // numPicBullet guard below. The document is genuinely broken, so this
        // still surfaces as a validation error (success:false) — just an
        // actionable one. Tracked so the NRE catch can suppress the bare message
        // when we've already explained the real cause.
        var orphanRefErrors = DetectOrphanedHeaderFooterReferences(clone);
        errors.AddRange(orphanRefErrors);
        // Cross-format OPC structural check: [Content_Types].xml that declares
        // parts only via <Override> and omits <Default Extension="rels"> is
        // OPC-legal and passes schema validation, but real Word/Excel/PowerPoint
        // refuse to open it ("file may be corrupt"). Read the manifest from the
        // original package on disk (the in-memory clone re-serializes a healthy
        // Content_Types, so the defect is only visible in the source file).
        errors.AddRange(DetectMissingDefaultRelsContentType(package, filePath));
        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        // BUG-R6-08: documents containing w:numPicBullet can trip an NRE
        // inside SDK validation when one of its child accessors hits a
        // null. Materialise per-error with try/catch so a single problem
        // entry doesn't bring the whole `validate` command down. Surface
        // the exception as a synthetic ValidationError instead of
        // bubbling out as a process-level crash.
        IEnumerable<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> raw;
        try
        {
            raw = validator.Validate(clone);
        }
        catch (Exception ex)
        {
            // BUG-R5B(BUG2): when the pre-results NRE is the orphaned-reference
            // one we already diagnosed, suppress the bare exception text — the
            // actionable error(s) are already in the list. Otherwise surface the
            // exception so unknown failures stay visible.
            if (!(ex is NullReferenceException && orphanRefErrors.Count > 0))
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"Validator threw before producing results: {ex.GetType().Name}: {ex.Message}",
                    null, null));
            }
            return errors;
        }
        // The IEnumerable is lazy — iterate with try/catch so one bad
        // error entry does not abort the rest.
        using var enumerator = raw.GetEnumerator();
        while (true)
        {
            DocumentFormat.OpenXml.Validation.ValidationErrorInfo? e = null;
            try
            {
                if (!enumerator.MoveNext()) break;
                e = enumerator.Current;
            }
            catch (NullReferenceException nre)
            {
                errors.Add(new ValidationError(
                    "ValidatorNullReference",
                    $"SDK validator hit a null while inspecting next error: {nre.Message}",
                    null, null));
                continue;
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"SDK validator threw while inspecting next error: {ex.GetType().Name}: {ex.Message}",
                    null, null));
                continue;
            }
            if (e == null) continue;
            try
            {
                errors.Add(new ValidationError(
                    e.ErrorType.ToString(),
                    e.Description,
                    e.Path?.XPath,
                    e.Part?.Uri.ToString()));
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"Failed to materialise validation error: {ex.GetType().Name}: {ex.Message}",
                    null, null));
            }
        }
        // Drop known SDK-validator false positives on chartEx (cx:) val-attribute
        // elements — see IsBenignChartExValAttributeError. A valid pareto /
        // histogram chart opens and renders correctly in real Excel; without this
        // it would fail `validate` (and the Delivery Gate) over a gap in the
        // OpenXmlValidator's schema model, not a real defect.
        errors.RemoveAll(IsBenignChartExValAttributeError);
        return errors;
    }

    /// <summary>
    /// True for the OpenXmlValidator false positives on chartEx (cx:) elements
    /// whose id is carried in a <c>val</c> ATTRIBUTE — <c>&lt;cx:axisId val="1"/&gt;</c>,
    /// <c>&lt;cx:binCount val="5"/&gt;</c>, <c>&lt;cx:binSize val="3"/&gt;</c>. That is the
    /// form every file real Excel writes and requires (the text-content form the
    /// SDK models instead makes Excel refuse the whole workbook, 0x800A03EC — see
    /// <c>MakeCxAxisId</c> in <c>Core/Chart/ChartExBuilder.cs</c>). The validator
    /// can't model the attribute form, so it emits "the 'val' attribute is not
    /// declared" and "the text value cannot be empty" on these elements. They are
    /// not real defects; suppress them narrowly (chartEx part + one of those three
    /// elements + one of those two messages) so genuine errors still surface.
    /// </summary>
    private static bool IsBenignChartExValAttributeError(ValidationError e)
    {
        var path = e.Path ?? "";
        var part = e.Part ?? "";
        bool inChartEx = part.Contains("extendedChart", StringComparison.OrdinalIgnoreCase)
            || part.Contains("chartEx", StringComparison.OrdinalIgnoreCase)
            || path.Contains("cx:", StringComparison.OrdinalIgnoreCase);
        if (!inChartEx) return false;
        bool valAttrElement = path.Contains(":axisId", StringComparison.OrdinalIgnoreCase)
            || path.Contains(":binCount", StringComparison.OrdinalIgnoreCase)
            || path.Contains(":binSize", StringComparison.OrdinalIgnoreCase);
        if (!valAttrElement) return false;
        var d = e.Description ?? "";
        return d.Contains("'val' attribute is not declared", StringComparison.OrdinalIgnoreCase)
            || d.Contains("text value cannot be empty", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// BUG-R5B(BUG2): detect header/footer references whose relationship id does
    /// not resolve to a part. Such a dangling <c>w:headerReference</c> /
    /// <c>w:footerReference</c> makes the OpenXML SDK validator throw a bare
    /// <see cref="NullReferenceException"/> before producing any results, so the
    /// caller would otherwise see "Validator threw … NullReferenceException"
    /// with no clue about the real problem. Return one actionable schema-style
    /// error per orphaned reference instead. Word-only; other formats have no
    /// equivalent reference shape (returns an empty list).
    /// </summary>
    private static List<ValidationError> DetectOrphanedHeaderFooterReferences(OpenXmlPackage package)
    {
        var result = new List<ValidationError>();
        if (package is not DocumentFormat.OpenXml.Packaging.WordprocessingDocument word)
            return result;
        var mainPart = word.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (mainPart == null || body == null) return result;

        // r:id values the MainDocumentPart actually knows about.
        var knownIds = new HashSet<string>(
            mainPart.Parts.Select(p => p.RelationshipId),
            StringComparer.Ordinal);
        // ExternalRelationships (rare for hdr/ftr, but be exhaustive).
        foreach (var rel in mainPart.ExternalRelationships)
            knownIds.Add(rel.Id);

        void Check(string id, string kind, string? typeName)
        {
            if (string.IsNullOrEmpty(id))
            {
                result.Add(new ValidationError(
                    "OrphanedReference",
                    $"{kind} reference (type '{typeName ?? "default"}') has no relationship id (r:id); "
                    + "the referenced part cannot be resolved.",
                    "/w:document/w:body", mainPart.Uri.ToString()));
                return;
            }
            if (!knownIds.Contains(id))
            {
                result.Add(new ValidationError(
                    "OrphanedReference",
                    $"{kind} reference '{id}' (type '{typeName ?? "default"}') has no matching relationship "
                    + "in the document part; the referenced part is missing. "
                    + $"Remove the dangling {kind.ToLowerInvariant()}Reference or add the {kind.ToLowerInvariant()} part it points at.",
                    "/w:document/w:body", mainPart.Uri.ToString()));
            }
        }

        // SectionProperties live both at body level (final section) and inside
        // each section-break paragraph's pPr (inline sectPr). Descendants covers
        // both without double-counting.
        foreach (var sectPr in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>())
        {
            foreach (var hr in sectPr.Elements<DocumentFormat.OpenXml.Wordprocessing.HeaderReference>())
                Check(hr.Id?.Value ?? "", "Header", hr.Type?.InnerText);
            foreach (var fr in sectPr.Elements<DocumentFormat.OpenXml.Wordprocessing.FooterReference>())
                Check(fr.Id?.Value ?? "", "Footer", fr.Type?.InnerText);
        }
        return result;
    }

    /// <summary>
    /// Cross-format OPC structural check: every OPC package contains relationship
    /// (<c>.rels</c>) parts, so its <c>[Content_Types].xml</c> manifest must
    /// declare their content type. The canonical, Word/Office-required way is a
    /// <c>&lt;Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/&gt;</c>
    /// entry. A manifest that lists parts only via <c>&lt;Override&gt;</c> and omits
    /// this <c>&lt;Default&gt;</c> is OPC-legal and passes schema validation, yet
    /// real Word/Excel/PowerPoint refuse to open the file ("file may be corrupt").
    /// Flag the missing <c>&lt;Default Extension="rels"&gt;</c> so callers learn the
    /// true reason an otherwise schema-clean file won't open.
    ///
    /// Reads the raw manifest from the original file on disk via
    /// <see cref="TryReadByZipUri"/> (the in-memory SDK clone re-serializes a
    /// healthy Content_Types, so the defect is invisible there). Returns an empty
    /// list when the manifest cannot be read (no file path / not on disk) — never
    /// a false positive.
    /// </summary>
    private static List<ValidationError> DetectMissingDefaultRelsContentType(
        OpenXmlPackage package, string? filePath)
    {
        var result = new List<ValidationError>();
        if (filePath == null) return result;

        string? manifest;
        try
        {
            manifest = TryReadByZipUri(package, filePath, "[Content_Types].xml");
        }
        catch
        {
            // Unreadable/empty manifest: leave to the other checks; don't guess.
            return result;
        }
        if (string.IsNullOrEmpty(manifest)) return result;

        System.Xml.Linq.XDocument doc;
        try
        {
            doc = System.Xml.Linq.XDocument.Parse(manifest);
        }
        catch
        {
            return result;
        }

        System.Xml.Linq.XNamespace ct =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        var hasDefaultRels = doc.Descendants(ct + "Default").Any(d =>
            string.Equals(
                (string?)d.Attribute("Extension"), "rels",
                StringComparison.OrdinalIgnoreCase));

        if (!hasDefaultRels)
        {
            result.Add(new ValidationError(
                "PackageStructure",
                "[Content_Types].xml is missing a <Default Extension=\"rels\" "
                + "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/> "
                + "declaration. Every OPC package contains .rels relationship parts, "
                + "and Word/Excel/PowerPoint refuse to open a file whose Content_Types "
                + "declares parts only via <Override> without this <Default> entry "
                + "(\"the file may be corrupt\"). Add the <Default Extension=\"rels\" ...> "
                + "entry to [Content_Types].xml to make the package openable.",
                null, "/[Content_Types].xml"));
        }
        return result;
    }

    // Cache the reflective lookup of OpenXmlPart's private loaded-root field.
    // null once resolution has been attempted-and-failed (older/newer SDK with a
    // renamed field) so we degrade to a no-op instead of throwing per call.
    private static System.Reflection.FieldInfo? _rootElementField;
    private static bool _rootElementFieldResolved;

    /// <summary>
    /// Flush each ALREADY-LOADED part's in-memory DOM back to its part stream so
    /// a subsequent Clone/GetStream of the LIVE package reads in-sync bytes and
    /// cannot desync the SDK's dirty-tracking (which would serialize the part
    /// EMPTY on the caller's next Save). Only loaded parts are touched: an
    /// unloaded part cannot be dirty, and force-loading one would make the
    /// caller's Save re-serialize an untouched part. Best-effort: any failure to
    /// peek/flush a part is swallowed — validate must never throw on a quirk here.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2075",
        Justification = "Reflects over OpenXmlPart's private _rootElement field; officecli ships framework-dependent (not trimmed/AOT). Best-effort with a null guard if the field is removed.")]
    private static void FlushLoadedPartRoots(OpenXmlPackage package)
    {
        if (!_rootElementFieldResolved)
        {
            _rootElementFieldResolved = true;
            // OpenXmlPart stores its loaded root in a private "_rootElement"
            // field (no public IsRootElementLoaded in DocumentFormat.OpenXml
            // 3.x). Walk the type hierarchy to find it.
            for (var t = typeof(OpenXmlPart); t != null; t = t.BaseType)
            {
                var f = t.GetField("_rootElement",
                    System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.NonPublic);
                if (f != null) { _rootElementField = f; break; }
            }
        }
        if (_rootElementField == null) return;   // SDK shape changed — no-op

        foreach (var part in package.GetAllParts())
        {
            try
            {
                if (_rootElementField.GetValue(part) is OpenXmlPartRootElement root)
                    root.Save();   // DOM -> part stream; idempotent on clean parts
            }
            catch
            {
                // Reflection/serialize hiccup on one part must not fail validate.
            }
        }
    }

    /// <summary>
    /// Walk every XML part in the package and try to parse it. OpenXmlValidator
    /// silently skips parts that can't be read, so a deck whose presentation.xml
    /// or any slide part is malformed XML (well-formed zip, broken XML inside)
    /// otherwise reports "validate passed: no errors found" — a dangerous
    /// false negative that hides obvious corruption. Surface a synthetic
    /// MalformedXml validation error for each part that fails to parse so the
    /// caller sees real failure instead of clean success.
    /// </summary>
    private static List<ValidationError> PreflightXmlParts(OpenXmlPackage package)
    {
        var errors = new List<ValidationError>();
        foreach (var part in package.GetAllParts())
        {
            // Only inspect XML parts; binary streams (images, fonts, embedded
            // OLE) don't participate in schema validation. The naive
            // ContentType.Contains("xml") screen matches every OOXML content
            // type via the "openxmlformats" substring (e.g.
            // "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            // for an OLE-embedded .docx), so a perfectly valid binary
            // OOXML embedding was mis-classified as malformed XML on every
            // validate run. RFC 7303 / OOXML restrict actual XML content types
            // to the "+xml" structured-suffix or the plain text/xml family;
            // gate the preflight on those instead.
            var ct = part.ContentType ?? "";
            if (!IsXmlContentType(ct)) continue;
            try
            {
                using var s = part.GetStream(FileMode.Open, FileAccess.Read);
                using var r = XmlReader.Create(s, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreWhitespace = false,
                });
                while (r.Read()) { /* drain */ }
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "MalformedXml",
                    $"Part contains malformed XML and was skipped by schema validation: {ex.GetType().Name}: {ex.Message}",
                    null,
                    part.Uri.ToString()));
            }
        }
        return errors;
    }

    private static bool IsXmlContentType(string ct)
    {
        if (string.IsNullOrEmpty(ct)) return false;
        // Drop parameters (e.g. "; charset=utf-8") and lowercase the type.
        var semi = ct.IndexOf(';');
        var bare = (semi >= 0 ? ct.Substring(0, semi) : ct).Trim().ToLowerInvariant();
        // Structured-suffix form: anything/anything+xml — covers every OOXML
        // schema part (...wordprocessingml.document.main+xml,
        // ...presentationml.slide+xml, ...drawingml.theme+xml, etc.).
        if (bare.EndsWith("+xml")) return true;
        // Generic XML content types.
        return bare == "application/xml" || bare == "text/xml";
    }

    private static void TryAddNamespace(XmlNamespaceManager nsManager, string prefix, string uri)
    {
        if (string.IsNullOrEmpty(nsManager.LookupNamespace(prefix)))
        {
            nsManager.AddNamespace(prefix, uri);
        }
    }

    // ==================== Zip-URI part lookup ====================
    //
    // Rule: any partPath ending in `.xml` is treated as a literal zip-internal
    // URI (e.g. `/xl/worksheets/sheet1.xml`, `/word/footnotes.xml`,
    // `/ppt/slides/slide1.xml`). We walk the entire part tree of the package
    // and match against `OpenXmlPart.Uri.OriginalString`.
    //
    // This supersedes the per-handler hand-curated alias tables, which could
    // never be complete (only covered global parts like /xl/workbook.xml).
    // Semantic paths (`/Sheet1`, `/workbook`, `/document`, `/header[1]`) still
    // route through the handler's own switch — only `.xml`-suffixed inputs
    // hit this lookup.

    /// <summary>
    /// Returns true if `partPath` should be resolved as a literal zip-internal
    /// URI rather than a semantic short name. Trims surrounding whitespace
    /// and discards any URI fragment (`#...`) or query (`?...`) suffix so
    /// `/xl/workbook.xml#frag` and `/xl/workbook.xml?x=1` both classify as
    /// zip-URI inputs (rather than silently falling through to the
    /// semantic-path "Available: ..." dispatcher).
    ///
    /// Accepts `.xml`, `.rels` (relationship parts), and the literal
    /// `[Content_Types].xml` package manifest. The first two are normal OPC
    /// parts; `[Content_Types].xml` is package metadata reachable only
    /// through a separate code path.
    /// </summary>
    public static bool IsZipUriPath(string partPath)
    {
        if (partPath == null) return false;
        var s = StripUriSuffixes(partPath.AsSpan().Trim());
        return s.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || s.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
    }

    private static ReadOnlySpan<char> StripUriSuffixes(ReadOnlySpan<char> s)
    {
        var cut = s.IndexOfAny('#', '?');
        return cut >= 0 ? s[..cut] : s;
    }

    /// <summary>
    /// Walk the entire part tree of a package and return the part whose
    /// `Uri.OriginalString` matches `partPath` (with leading-slash
    /// normalization). Returns null if no part matches.
    /// </summary>
    /// <summary>
    /// Resolve a zip-URI path through the SDK's own underlying IPackage.
    /// Used as the primary fallback after the typed-OpenXmlPart graph,
    /// because it shares the SDK's file handle and so works correctly when
    /// the file is held open editable (resident mode) where a fresh BCL
    /// Package.Open would fail with a FileShare conflict. Returns null if
    /// the part does not exist.
    /// </summary>
    private static string? TryReadViaSdkPackage(OpenXmlPackage package, string partPath)
    {
        try
        {
            var clean = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
            var target = clean.StartsWith('/') ? clean : "/" + clean;
            Uri uri;
            try { uri = new Uri(target, UriKind.Relative); } catch { return null; }

#pragma warning disable OOXML0001
            var pkg = package.GetPackage();
#pragma warning restore OOXML0001
            if (!pkg.PartExists(uri)) return null;
            var part = pkg.GetPart(uri);

            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();
            var stripped = StripXmlProlog(content);
            if (stripped.Length == 0)
                throw new InvalidDataException(
                    $"Part '{target}' contains no root element (only an XML " +
                    $"declaration, whitespace, or BOM).");
            return stripped;
        }
        catch (InvalidDataException) { throw; }
        catch
        {
            return null;
        }
    }

    public static OpenXmlPart? FindPartByZipUri(OpenXmlPackage package, string partPath)
    {
        // Trim surrounding whitespace, discard fragment/query, and normalize
        // leading slash. Fragments and query strings are not part of OPC
        // URIs; users may inadvertently type them and we should resolve
        // against the bare part path.
        partPath = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
        var target = partPath.StartsWith('/') ? partPath : "/" + partPath;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        return Walk(package);

        OpenXmlPart? Walk(OpenXmlPartContainer container)
        {
            foreach (var rel in container.Parts)
            {
                var p = rel.OpenXmlPart;
                var uri = p.Uri.OriginalString;
                if (!seen.Add(uri)) continue;
                if (string.Equals(uri, target, StringComparison.OrdinalIgnoreCase))
                    return p;
                var nested = Walk(p);
                if (nested != null) return nested;
            }
            return null;
        }
    }

    /// <summary>
    /// Read a part's XML content as a string. Prefer the typed
    /// <see cref="OpenXmlPartRootElement"/> when available (preserves
    /// canonical SDK serialization); fall back to the underlying stream for
    /// untyped XML parts (e.g. CustomXml).
    ///
    /// Output omits the &lt;?xml ?&gt; prolog uniformly so that:
    ///   raw /workbook              (semantic path, typed OuterXml)
    ///   raw /xl/workbook.xml       (zip URI, typed OuterXml)
    ///   raw /customXml/item1.xml   (zip URI, untyped stream)
    /// all produce element-only output. The semantic short-name path has
    /// always done this; this method extends the convention to untyped
    /// parts so zip-URI calls don't randomly include the prolog depending
    /// on whether the SDK strongly-typed the target part.
    /// </summary>
    /// <summary>
    /// Resolve a zip-URI path to its content. Tries the OpenXmlPart graph
    /// first (typed parts — preserves SDK-canonical serialization for parts
    /// that have a strongly-typed root); falls back to the underlying OPC
    /// package (covers relationship parts `.rels` and any XML part the
    /// SDK doesn't surface as a typed OpenXmlPart).
    ///
    /// Returns null if no part matches; throws InvalidDataException if the
    /// part exists but contains no root element.
    /// </summary>
    public static string? TryReadByZipUri(OpenXmlPackage package, string? filePath, string partPath)
    {
        // Typed-part path first (preserves SDK-canonical serialization for
        // strongly-typed parts).
        var typed = FindPartByZipUri(package, partPath);
        if (typed != null) return ReadPartXml(typed);

        // Then: SDK's own underlying IPackage via reflection. This sees
        // every .rels part the SDK is managing AND coexists with the SDK
        // file handle (no second-handle FileShare conflict — important for
        // resident mode where the file is open editable and a fresh
        // BCL Package.Open would fail).
        var sdkResult = TryReadViaSdkPackage(package, partPath);
        if (sdkResult != null) return sdkResult;

        if (filePath == null) return null;

        // Special case: `[Content_Types].xml` is the OPC package manifest,
        // not a part. System.IO.Packaging.Package does not expose it; read
        // it as a literal zip entry.
        var trimmed = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
        if (trimmed.TrimStart('/').Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
                var entry = zip.Entries.FirstOrDefault(e =>
                    e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));
                if (entry == null) return null;
                using var es = entry.Open();
                using var er = new StreamReader(es);
                var ec = er.ReadToEnd();
                var es2 = StripXmlProlog(ec);
                return es2.Length == 0
                    ? throw new InvalidDataException("[Content_Types].xml is empty.")
                    : es2;
            }
            catch (Exception ex) when (ex is not InvalidDataException)
            {
                return null;
            }
        }

        var clean = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
        var target = clean.StartsWith('/') ? clean : "/" + clean;
        Uri uri;
        try { uri = new Uri(target, UriKind.Relative); } catch { return null; }

        Package bclPkg;
        try
        {
            // FileShare.ReadWrite so we coexist with the SDK's existing handle.
            // Package internally opens the underlying stream with
            // FileAccess.Read here.
            bclPkg = Package.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
        catch
        {
            // SDK opened with FileShare.None (e.g. editable mode on Windows),
            // or the file is gone — give up cleanly.
            return null;
        }

        try
        {
            if (!bclPkg.PartExists(uri)) return null;
            var part = bclPkg.GetPart(uri);
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();
            var stripped = StripXmlProlog(content);
            if (stripped.Length == 0)
                throw new InvalidDataException(
                    $"Part '{target}' contains no root element (only an XML " +
                    $"declaration, whitespace, or BOM).");
            return stripped;
        }
        finally
        {
            bclPkg.Close();
        }
    }

    public static string ReadPartXml(OpenXmlPart part)
    {
        if (part.RootElement is OpenXmlPartRootElement root && root != null)
            return root.OuterXml;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        var content = reader.ReadToEnd();
        var stripped = StripXmlProlog(content);
        if (stripped.Length == 0)
            throw new InvalidDataException(
                $"Part '{part.Uri.OriginalString}' contains no root element " +
                $"(only an XML declaration, whitespace, or BOM). The package " +
                $"may be corrupt; investigate before treating output as data.");
        return stripped;
    }

    private static string StripXmlProlog(string xml)
    {
        var s = xml.AsSpan().TrimStart();
        // Loop: handle multiple stacked prologs / BOMs (defensive — input may
        // be byte-concatenated from upstream tools or a corrupted package).
        while (s.Length > 0)
        {
            // BOM (U+FEFF). StreamReader normally consumes it but we may be
            // reading a re-encoded inner segment.
            if (s[0] == '﻿') { s = s[1..].TrimStart(); continue; }

            // XML declaration: per spec must be `<?xml` followed by whitespace
            // or `?>`. Crucially must NOT match other PIs whose target starts
            // with `xml` (e.g. `<?xml-stylesheet ...?>`), which is a legal
            // processing instruction we must preserve.
            if (s.Length >= 6
                && s[0] == '<' && s[1] == '?'
                && s[2] == 'x' && s[3] == 'm' && s[4] == 'l'
                && (s[5] == ' ' || s[5] == '\t' || s[5] == '\n' || s[5] == '\r'
                    || (s[5] == '?' && s.Length >= 7 && s[6] == '>')))
            {
                var end = s.IndexOf("?>", StringComparison.Ordinal);
                if (end < 0) break;
                s = s[(end + 2)..].TrimStart();
                continue;
            }
            break;
        }
        return s.ToString();
    }

    /// <summary>
    /// Write XML content into a part's stream, replacing prior contents.
    /// </summary>
    public static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }
}
