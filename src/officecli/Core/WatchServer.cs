// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// CONSISTENCY(watch-isolation): 本文件不引用 OfficeCli.Handlers,不打开文件,不写盘。
// 见 CLAUDE.md "Watch Server Rules"。要放宽这条红线,
// grep "CONSISTENCY(watch-isolation)" 找全 watch 子系统所有文件项目级一起评审。

using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Pure SSE relay server. Never opens the document file.
/// Receives pre-rendered HTML from command processes via named pipe,
/// forwards to browsers via SSE.
/// </summary>
public class WatchServer : IDisposable
{
    private readonly string _filePath;
    private readonly string _pipeName;
    private readonly int _port;
    private readonly TcpListener _tcpListener;
    private readonly List<NetworkStream> _sseClients = new();
    private readonly object _sseLock = new();
    private CancellationTokenSource _cts = new();
    private string _currentHtml = "";
    private int _version = 0;
    private bool _disposed;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private readonly TimeSpan _idleTimeout;

    // Current selection — paths of elements selected in any connected browser.
    // Single shared list (last-write-wins): all browsers viewing the same file see
    // the same selection. CLI reads this via the named pipe "get-selection" command.
    //
    // CONSISTENCY(path-stability): selection 和 mark 共享同一套裸位置寻址契约,
    // 没有指纹/漂移检测。要升级到稳定 ID,grep "CONSISTENCY(path-stability)"
    // 找全所有 deferred 站点项目级一起改。见 CLAUDE.md "Design Principles"。
    private List<string> _currentSelection = new();
    private readonly object _selectionLock = new();

    // Current marks — advisory annotations attached to document paths. Live in
    // memory only. Server never opens the document and never inspects DOM —
    // marks are pure metadata; the browser computes match positions client-side.
    //
    // CONSISTENCY(path-stability): 元素删除/位置漂移的处理刻意和 selection 一致 ——
    // 裸位置寻址,无指纹,无漂移检测。stale 仅在 path 解析失败或 find 不命中时由
    // 客户端报告设置。见 CLAUDE.md "Design Principles" + "Watch Server Rules"。
    // 要修复成稳定 ID 路径,grep "CONSISTENCY(path-stability)" 找全所有 deferred 站点
    // (selection / mark / 未来其它 path 消费者)项目级一起改,不要在 mark 单点改。
    private readonly List<WatchMark> _currentMarks = new();
    private readonly object _marksLock = new();
    private int _marksVersion = 0;
    private int _nextMarkId = 1;

    private const string WaitingHtml = """
        <html><head><meta charset="utf-8"><title>Watching...</title>
        <style>body{font-family:system-ui;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f5f5f5;color:#666;}
        .msg{text-align:center;}</style></head>
        <body><div class="msg"><h2>Waiting for first update...</h2><p>Run an officecli command to see the preview.</p></div></body></html>
        """;

    private const string SseScript = """
        <script>
        (function() {
            var es = new EventSource('/events');
            var _scrollTimer = null;
            // ===== Selection sync =====
            // Single source of truth: server's currentSelection. We keep a local
            // mirror updated by the server's SSE 'selection-update' broadcasts so
            // that we can re-apply highlights after every DOM swap.
            var _selection = [];
            function applySelectionToDom() {
                document.querySelectorAll('.officecli-selected').forEach(function(el) {
                    el.classList.remove('officecli-selected');
                });
                _selection.forEach(function(path) {
                    try {
                        var sel = '[data-path="' + path.replace(/"/g, '\\"') + '"]';
                        document.querySelectorAll(sel).forEach(function(el) {
                            el.classList.add('officecli-selected');
                        });
                    } catch (e) {}
                });
            }
            function postSelection(paths) {
                fetch('/api/selection', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ paths: paths })
                }).catch(function() {});
            }
            // Inject selection + mark highlight CSS
            (function() {
                var style = document.createElement('style');
                style.textContent =
                    '.officecli-selected{outline:2px solid #2196f3 !important;' +
                    'outline-offset:2px;' +
                    'box-shadow:0 0 12px rgba(33,150,243,0.6) !important;' +
                    'z-index:1000;}' +
                    '.officecli-mark{background:#ffeb3b;border-radius:2px;padding:0 1px;}' +
                    '.officecli-mark-block{outline:2px dashed #ffc107;outline-offset:2px;}' +
                    '.officecli-mark-stale{background:#e0e0e0 !important;opacity:0.55;text-decoration:line-through;}';
                document.head.appendChild(style);
            })();
            // ===== Marks =====
            // Server is the source of truth. The browser mirrors _marks via SSE
            // 'mark-update' broadcasts and re-applies them after every DOM swap.
            //
            // CONSISTENCY(find-regex): literal vs regex detection uses the r"..." /
            // r'...' raw-string prefix rule from WordHandler.Set.cs:60-61. If that
            // protocol changes, grep "CONSISTENCY(find-regex)" and update every site
            // (set handler, mark CLI, server, this JS) together. Do NOT diverge here.
            //
            // CONSISTENCY(path-stability): when a mark's path no longer resolves or
            // its find no longer matches, we flip a visual-only stale class and
            // move on — same naive positional model as selection. No fingerprint,
            // no drift detection. grep "CONSISTENCY(path-stability)" for deferred
            // sites. See CLAUDE.md Watch Server Rules.
            var _marks = [];
            function _isRegexFind(find) {
                if (!find || find.length < 3) return false;
                return (find.charAt(0) === 'r' &&
                        (find.charAt(1) === '"' || find.charAt(1) === "'") &&
                        find.charAt(find.length - 1) === find.charAt(1));
            }
            function _extractRegexPattern(find) {
                // r"..." or r'...' — strip the 2-char prefix and 1-char suffix
                return find.substring(2, find.length - 1);
            }
            function _normalizeNfc(s) {
                try { return s.normalize('NFC'); } catch (e) { return s; }
            }
            function _markTitle(m) {
                var find = m.find || '';
                var tofix = m.tofix || '';
                var note = m.note || '';
                if (tofix) {
                    var head = find ? (find + ' → ' + tofix) : ('→ ' + tofix);
                    return note ? (head + '\n' + note) : head;
                }
                return note;
            }
            function _clearMarks() {
                // Unwrap every existing .officecli-mark span, restoring original text
                // nodes. Iterate a snapshot because replaceWith mutates the NodeList.
                var spans = Array.prototype.slice.call(
                    document.querySelectorAll('.officecli-mark'));
                for (var i = 0; i < spans.length; i++) {
                    var sp = spans[i];
                    var parent = sp.parentNode;
                    if (!parent) continue;
                    while (sp.firstChild) parent.insertBefore(sp.firstChild, sp);
                    parent.removeChild(sp);
                    // Merge adjacent text nodes so future indexOf calls span the whole run
                    parent.normalize();
                }
                // Drop block-mark outlines and any stale inline overrides
                var blocks = Array.prototype.slice.call(
                    document.querySelectorAll('.officecli-mark-block'));
                for (var j = 0; j < blocks.length; j++) {
                    blocks[j].classList.remove('officecli-mark-block');
                    blocks[j].classList.remove('officecli-mark-stale');
                    if (blocks[j].dataset && blocks[j].dataset.officecliMarkBg) {
                        blocks[j].style.backgroundColor = '';
                        delete blocks[j].dataset.officecliMarkBg;
                    }
                }
            }
            // Walk the element's text nodes and return
            //   { text: concatenated NFC text, map: [ {node, start, end} ... ] }
            // so we can map absolute char offsets in `text` back to specific text nodes.
            function _buildTextMap(el) {
                var walker = document.createTreeWalker(
                    el, NodeFilter.SHOW_TEXT, null, false);
                var parts = [];
                var map = [];
                var cursor = 0;
                var n;
                while ((n = walker.nextNode())) {
                    var v = _normalizeNfc(n.nodeValue || '');
                    if (v.length === 0) continue;
                    parts.push(v);
                    map.push({ node: n, start: cursor, end: cursor + v.length });
                    cursor += v.length;
                }
                return { text: parts.join(''), map: map };
            }
            function _findNodeAt(map, offset) {
                // Linear scan — element text count is small; binary search unnecessary.
                for (var i = 0; i < map.length; i++) {
                    if (offset >= map[i].start && offset < map[i].end) {
                        return { node: map[i].node, local: offset - map[i].start };
                    }
                }
                // Offset at very end of last node
                if (map.length > 0 && offset === map[map.length - 1].end) {
                    var last = map[map.length - 1];
                    return { node: last.node, local: last.end - last.start };
                }
                return null;
            }
            function _wrapRange(el, startOff, endOff, map, markId, color, title, stale) {
                var s = _findNodeAt(map, startOff);
                var e = _findNodeAt(map, endOff);
                if (!s || !e) return false;
                var range = document.createRange();
                try {
                    range.setStart(s.node, s.local);
                    range.setEnd(e.node, e.local);
                } catch (err) {
                    return false;
                }
                var span = document.createElement('span');
                span.className = stale ? 'officecli-mark officecli-mark-stale' : 'officecli-mark';
                span.setAttribute('data-mark-id', markId);
                if (color) span.style.backgroundColor = color;
                if (title) span.title = title;
                try {
                    range.surroundContents(span);
                } catch (err) {
                    // surroundContents throws if the range spans a non-Text boundary.
                    // Fallback: extract + insert. Loses the "single wrapper" property but
                    // still applies visual styling to the content (the span wraps a
                    // DocumentFragment which still carries the class).
                    try {
                        var frag = range.extractContents();
                        span.appendChild(frag);
                        range.insertNode(span);
                    } catch (err2) {
                        return false;
                    }
                }
                return true;
            }
            function applyMarks() {
                _clearMarks();
                if (!_marks || _marks.length === 0) return;
                for (var mi = 0; mi < _marks.length; mi++) {
                    var m = _marks[mi];
                    if (!m || !m.path) continue;
                    var el;
                    try {
                        var sel = '[data-path="' + m.path.replace(/"/g, '\\"') + '"]';
                        el = document.querySelector(sel);
                    } catch (e) { el = null; }
                    if (!el) {
                        // CONSISTENCY(path-stability): path no longer resolves — skip.
                        // No drift detection, no fallback lookup. Consistent with selection.
                        continue;
                    }
                    var title = _markTitle(m);
                    var color = m.color || '';
                    // No find → the whole element is the mark
                    if (!m.find) {
                        el.classList.add('officecli-mark-block');
                        if (m.stale) el.classList.add('officecli-mark-stale');
                        if (title) el.title = title;
                        if (color) {
                            el.style.backgroundColor = color;
                            if (!el.dataset) el.dataset = {};
                            el.dataset.officecliMarkBg = '1';
                        }
                        continue;
                    }
                    // Find has a value → locate matches and wrap each.
                    // CONSISTENCY(find-regex): detect r"..." / r'...' prefix the same way
                    // the C# side does (see WordHandler.Set.cs:60-61 and
                    // CommandBuilder.Mark.cs). Keep these in sync.
                    var tm = _buildTextMap(el);
                    var text = tm.text;
                    if (text.length === 0) continue;
                    var hitCount = 0;
                    if (_isRegexFind(m.find)) {
                        var patt = _extractRegexPattern(m.find);
                        var re;
                        try { re = new RegExp(patt, 'g'); }
                        catch (rxErr) { continue; }
                        // Re-read tm after each successful wrap — wrapping mutates
                        // the DOM, invalidating text node references. Start over
                        // from the remaining tail text.
                        var cursor = 0;
                        while (true) {
                            re.lastIndex = cursor;
                            var mr = re.exec(text);
                            if (!mr) break;
                            var mStart = mr.index;
                            var mEnd = mr.index + mr[0].length;
                            if (mEnd === mStart) {
                                // Zero-width match — advance to avoid infinite loop
                                cursor = mEnd + 1;
                                if (cursor > text.length) break;
                                continue;
                            }
                            var freshMap = _buildTextMap(el);
                            if (_wrapRange(el, mStart, mEnd, freshMap.map,
                                           m.id, color, title, m.stale)) {
                                hitCount++;
                            }
                            // After a wrap the text content is unchanged (we only
                            // insert a span, the text characters stay in place), so
                            // we can keep matching in the same `text` string.
                            cursor = mEnd;
                            if (hitCount > 500) break; // safety cap
                        }
                    } else {
                        var needle = _normalizeNfc(m.find);
                        if (needle.length === 0) continue;
                        var from = 0;
                        while (true) {
                            var idx = text.indexOf(needle, from);
                            if (idx < 0) break;
                            var fm = _buildTextMap(el);
                            if (_wrapRange(el, idx, idx + needle.length, fm.map,
                                           m.id, color, title, m.stale)) {
                                hitCount++;
                            }
                            from = idx + needle.length;
                            if (hitCount > 500) break;
                        }
                    }
                    if (hitCount === 0) {
                        // find supplied but nothing matched — visually mark the block
                        // as stale so the user can see the mark is "orphaned".
                        el.classList.add('officecli-mark-block');
                        el.classList.add('officecli-mark-stale');
                        if (title) el.title = title;
                    }
                }
            }
            // Unified reapply hook used by every code path that swaps or mutates DOM.
            function reapplyDecorations() {
                applySelectionToDom();
                applyMarks();
            }
            window._officecliReapplyDecorations = reapplyDecorations;
            window._officecliApplyMarks = applyMarks;
            window._officecliSetMarks = function(arr) { _marks = arr || []; applyMarks(); };
            window._officecliGetMarks = function() { return _marks; };
            // Click handler — selects the closest element with [data-path].
            // shift/ctrl/cmd toggle multi-select; plain click replaces.
            // Skipped if a rubber-band drag just finished.
            var _suppressNextClick = false;
            document.addEventListener('click', function(e) {
                if (_suppressNextClick) { _suppressNextClick = false; return; }
                var target = e.target.closest('[data-path]');
                if (!target) {
                    if (!e.shiftKey && !e.ctrlKey && !e.metaKey && _selection.length > 0) {
                        _selection = [];
                        postSelection([]);
                    }
                    return;
                }
                var path = target.getAttribute('data-path');
                if (!path) return;
                if (e.shiftKey || e.ctrlKey || e.metaKey) {
                    var idx = _selection.indexOf(path);
                    if (idx >= 0) _selection.splice(idx, 1);
                    else _selection.push(path);
                } else {
                    _selection = [path];
                }
                postSelection(_selection);
                e.preventDefault();
                e.stopPropagation();
            }, true);
            // ===== Rubber-band (box) selection =====
            // Press on empty space (no [data-path] under cursor) and drag to draw a
            // selection rectangle. Any element whose bounding box intersects the
            // rectangle gets selected. Shift adds to current selection; plain replaces.
            // Esc cancels mid-drag.
            var _rubber = null; // {startX, startY, shift, div}
            var _RUBBER_THRESHOLD = 5; // px before treating as drag (vs click)
            document.addEventListener('mousedown', function(e) {
                if (e.button !== 0) return;
                if (e.target.closest('[data-path]')) return;
                // Ignore mousedown inside scrollbars / sidebar / interactive UI
                if (e.target.closest('.sidebar, .sidebar-toggle, .page-counter, button, input, a')) return;
                _rubber = { startX: e.clientX, startY: e.clientY, shift: e.shiftKey, div: null };
            }, true);
            document.addEventListener('mousemove', function(e) {
                if (!_rubber) return;
                var dx = e.clientX - _rubber.startX;
                var dy = e.clientY - _rubber.startY;
                if (!_rubber.div) {
                    if (Math.abs(dx) < _RUBBER_THRESHOLD && Math.abs(dy) < _RUBBER_THRESHOLD) return;
                    var d = document.createElement('div');
                    d.id = '_officecli_rubber';
                    d.style.cssText = 'position:fixed;border:1.5px dashed #2196f3;' +
                        'background:rgba(33,150,243,0.12);pointer-events:none;' +
                        'z-index:99999;left:0;top:0;width:0;height:0;';
                    document.body.appendChild(d);
                    _rubber.div = d;
                }
                var x = Math.min(e.clientX, _rubber.startX);
                var y = Math.min(e.clientY, _rubber.startY);
                _rubber.div.style.left = x + 'px';
                _rubber.div.style.top = y + 'px';
                _rubber.div.style.width = Math.abs(dx) + 'px';
                _rubber.div.style.height = Math.abs(dy) + 'px';
            }, true);
            document.addEventListener('mouseup', function(e) {
                if (!_rubber) return;
                var rb = _rubber;
                _rubber = null;
                if (!rb.div) return; // didn't move enough — let normal click flow run
                rb.div.remove();
                var rect = {
                    left: Math.min(e.clientX, rb.startX),
                    top: Math.min(e.clientY, rb.startY),
                    right: Math.max(e.clientX, rb.startX),
                    bottom: Math.max(e.clientY, rb.startY)
                };
                // Hit-test: any [data-path] element that intersects the rect (counts
                // even partial overlap, like Figma — easier to use than full-contain)
                var hits = [];
                document.querySelectorAll('[data-path]').forEach(function(el) {
                    var r = el.getBoundingClientRect();
                    if (r.width === 0 || r.height === 0) return;
                    if (r.left < rect.right && r.right > rect.left &&
                        r.top < rect.bottom && r.bottom > rect.top) {
                        var p = el.getAttribute('data-path');
                        if (p && hits.indexOf(p) < 0) hits.push(p);
                    }
                });
                if (rb.shift) {
                    hits.forEach(function(p) {
                        if (_selection.indexOf(p) < 0) _selection.push(p);
                    });
                } else {
                    _selection = hits;
                }
                postSelection(_selection);
                // Suppress the synthetic click that fires right after mouseup, otherwise
                // the click-on-empty-space handler would clear the selection we just made.
                _suppressNextClick = true;
                e.preventDefault();
                e.stopPropagation();
            }, true);
            function _cancelRubber() {
                if (!_rubber) return;
                if (_rubber.div) _rubber.div.remove();
                _rubber = null;
            }
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') _cancelRubber();
            });
            // If the user alt-tabs / window loses focus mid-drag, the OS-level
            // mouseup never reaches us. Clean up so the rubber-band overlay
            // doesn't get stuck on screen and click handling stays sane.
            window.addEventListener('blur', _cancelRubber);
            document.addEventListener('visibilitychange', function() {
                if (document.hidden) _cancelRubber();
            });
            // Belt-and-suspenders: if a mouseup never came after a long enough
            // mousemove pause, drop the rubber-band on the next mouse re-entry.
            document.addEventListener('mouseleave', function(e) {
                // Only cancel if cursor truly left the page (relatedTarget == null)
                if (!e.relatedTarget && _rubber) _cancelRubber();
            });
            // SSE: receive selection and mark updates from the server.
            // This listener handles the lightweight metadata events; the heavier
            // DOM-swap events (full / replace / word-patch / ...) are handled by a
            // second listener registered further down.
            es.addEventListener('update', function(e) {
                var msg;
                try { msg = JSON.parse(e.data); } catch (err) { return; }
                if (msg.action === 'selection-update') {
                    _selection = msg.paths || [];
                    applySelectionToDom();
                } else if (msg.action === 'mark-update') {
                    // Monotonic version: clients may CAS on this value to skip
                    // redundant updates if they missed nothing. We just refresh.
                    _marks = msg.marks || [];
                    applyMarks();
                }
            });
            function scrollToSlide(num) {
                clearTimeout(_scrollTimer);
                _scrollTimer = setTimeout(function() {
                    var target = document.querySelector('.slide-container[data-slide="' + num + '"]');
                    if (target) target.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }, 300);
            }
            function syncThumbs() {
                var sidebar = document.querySelector('.sidebar');
                if (!sidebar) return;
                var slides = document.querySelectorAll('.main > .slide-container');
                var thumbs = sidebar.querySelectorAll('.thumb');
                // Remove extra thumbs
                for (var i = thumbs.length - 1; i >= slides.length; i--) {
                    thumbs[i].remove();
                }
                // Add missing thumbs
                for (var i = thumbs.length; i < slides.length; i++) {
                    var thumb = document.createElement('div');
                    thumb.className = 'thumb';
                    thumb.setAttribute('data-slide', i + 1);
                    thumb.innerHTML = '<div class="thumb-inner"></div><span class="thumb-num">' + (i + 1) + '</span>';
                    sidebar.appendChild(thumb);
                }
                // Renumber all thumbs
                sidebar.querySelectorAll('.thumb').forEach(function(t, i) {
                    t.setAttribute('data-slide', i + 1);
                    var num = t.querySelector('.thumb-num');
                    if (num) num.textContent = i + 1;
                });
                // Clear all thumb clones so buildThumbs re-creates them fresh
                sidebar.querySelectorAll('.thumb-inner').forEach(function(inner) {
                    var old = inner.querySelector('.thumb-slide');
                    if (old) old.remove();
                });
                if (typeof buildThumbs === 'function') buildThumbs();
                // Update page counter
                var counter = document.querySelector('.page-counter');
                if (counter) counter.textContent = '1 / ' + slides.length;
            }
            // Word diff-update: de-paginate, diff children, re-paginate (no full innerHTML swap)
            function wordDiffUpdate(msg) {
                var visiblePageNum = 0;
                document.querySelectorAll('.page-wrapper').forEach(function(w) {
                    var rect = w.getBoundingClientRect();
                    if (rect.top < window.innerHeight / 2) {
                        var p = w.querySelector('.page');
                        if (p) visiblePageNum = parseInt(p.getAttribute('data-page')) || 0;
                    }
                });
                fetch('/').then(function(r) { return r.text(); }).then(function(html) {
                    var doc = new DOMParser().parseFromString(html, 'text/html');
                    // Update styles
                    var oldStyles = document.querySelectorAll('head style');
                    var newStyles = doc.querySelectorAll('head style');
                    oldStyles.forEach(function(s) { s.remove(); });
                    newStyles.forEach(function(s) { document.head.appendChild(s.cloneNode(true)); });
                    // De-paginate: merge pagination-created pages back into section wrappers
                    var allW = Array.from(document.querySelectorAll('.page-wrapper'));
                    var curSec = null;
                    allW.forEach(function(w) {
                        if (w.hasAttribute('data-section')) { curSec = w; return; }
                        if (!curSec) return;
                        var src = w.querySelector('.page-body');
                        var dst = curSec.querySelector('.page-body');
                        if (src && dst) {
                            Array.from(src.children).forEach(function(c) {
                                if (!c.classList.contains('footnotes')) dst.appendChild(c);
                            });
                        }
                        w.remove();
                    });
                    // Diff per section
                    var contentAdded = false;
                    var oldSecs = Array.from(document.querySelectorAll('.page-wrapper[data-section]'));
                    var newSecs = Array.from(doc.querySelectorAll('.page-wrapper[data-section]'));
                    var maxS = Math.max(oldSecs.length, newSecs.length);
                    for (var si = 0; si < maxS; si++) {
                        if (si >= oldSecs.length) {
                            // New section added
                            var last = document.querySelector('.page-wrapper[data-section]:last-of-type');
                            if (last) last.after(newSecs[si].cloneNode(true));
                            continue;
                        }
                        if (si >= newSecs.length) { oldSecs[si].remove(); continue; }
                        var oldB = oldSecs[si].querySelector('.page-body');
                        var newB = newSecs[si].querySelector('.page-body');
                        if (!oldB || !newB) continue;
                        var oldK = Array.from(oldB.children).filter(function(c){ return !c.classList.contains('footnotes'); });
                        var newK = Array.from(newB.children).filter(function(c){ return !c.classList.contains('footnotes'); });
                        // Common prefix
                        var pi = 0;
                        while (pi < oldK.length && pi < newK.length && oldK[pi].outerHTML === newK[pi].outerHTML) pi++;
                        if (pi === oldK.length && pi === newK.length) continue; // identical
                        // Common suffix
                        var oi = oldK.length - 1, ni = newK.length - 1;
                        while (oi >= pi && ni >= pi && oldK[oi].outerHTML === newK[ni].outerHTML) { oi--; ni--; }
                        // Remove old diff range
                        for (var j = oi; j >= pi; j--) oldK[j].remove();
                        // Insert new diff range
                        var before = (oi + 1 < oldK.length) ? oldK[oi + 1] : oldB.querySelector('.footnotes');
                        for (var j = pi; j <= ni; j++) oldB.insertBefore(newK[j].cloneNode(true), before);
                        if (newK.length > oldK.length) contentAdded = true;
                    }
                    // Set scroll target
                    if (contentAdded) {
                        window._pendingScrollTo = '_last_page';
                    } else if (msg.scrollTo) {
                        window._pendingScrollTo = msg.scrollTo;
                    } else if (visiblePageNum > 0) {
                        window._pendingScrollTo = '.page[data-page="' + visiblePageNum + '"]';
                        window._pendingScrollBehavior = 'instant';
                    }
                    // Re-paginate (will also re-scale and remove freeze)
                    if (typeof window._wordPaginate === 'function') window._wordPaginate();
                    else { var f=document.getElementById('_sse_freeze'); if(f)f.remove(); }
                    // Re-apply selection + marks after DOM swap. Pagination is
                    // scheduled via setTimeout inside _wordPaginate; applyMarks
                    // walks the current DOM which already has the new content.
                    reapplyDecorations();
                });
            }
            // Track version for gap detection
            var _clientVersion = 0;
            // Apply server-side block patches directly to DOM
            function wordPatchUpdate(msg) {
                // De-paginate: merge pagination-created pages back into section wrappers
                var allW = Array.from(document.querySelectorAll('.page-wrapper'));
                var curSec = null;
                allW.forEach(function(w) {
                    if (w.hasAttribute('data-section')) { curSec = w; return; }
                    if (!curSec) return;
                    var src = w.querySelector('.page-body');
                    var dst = curSec.querySelector('.page-body');
                    if (src && dst) {
                        Array.from(src.children).forEach(function(c) {
                            if (!c.classList.contains('footnotes')) dst.appendChild(c);
                        });
                    }
                    w.remove();
                });
                var contentAdded = false;
                msg.patches.forEach(function(patch) {
                    if (patch.op === 'style') {
                        // Update CSS styles in head
                        document.querySelectorAll('head style').forEach(function(s) { s.remove(); });
                        var tmp = document.createElement('div');
                        tmp.innerHTML = patch.html;
                        tmp.querySelectorAll('style').forEach(function(s) { document.head.appendChild(s); });
                        return;
                    }
                    var bStart = document.querySelector('.wb[data-block="' + patch.block + '"]');
                    var bEnd = document.querySelector('.we[data-block="' + patch.block + '"]');
                    if (patch.op === 'remove') {
                        if (bStart && bEnd) {
                            // Remove everything between bStart and bEnd (inclusive)
                            var cur = bStart.nextSibling;
                            while (cur && cur !== bEnd) { var nx = cur.nextSibling; cur.remove(); cur = nx; }
                            bEnd.remove();
                            bStart.remove();
                        }
                    } else if (patch.op === 'replace') {
                        if (bStart && bEnd) {
                            // Remove old content between markers
                            var cur = bStart.nextSibling;
                            while (cur && cur !== bEnd) { var nx = cur.nextSibling; cur.remove(); cur = nx; }
                            // Insert new content before bEnd
                            var tmp = document.createElement('div');
                            tmp.innerHTML = patch.html;
                            while (tmp.firstChild) bEnd.parentNode.insertBefore(tmp.firstChild, bEnd);
                        }
                    } else if (patch.op === 'add') {
                        contentAdded = true;
                        var tmp = document.createElement('div');
                        tmp.innerHTML = '<span class="wb" data-block="' + patch.block + '" style="display:none"></span>' +
                            patch.html +
                            '<span class="we" data-block="' + patch.block + '" style="display:none"></span>';
                        // Find insertion point: after previous block's end, or before next block's begin
                        var prevEnd = patch.block > 1 ? document.querySelector('.we[data-block="' + (patch.block - 1) + '"]') : null;
                        if (prevEnd) {
                            var ref = prevEnd.nextSibling;
                            while (tmp.firstChild) prevEnd.parentNode.insertBefore(tmp.firstChild, ref);
                        } else {
                            var nextBegin = document.querySelector('.wb[data-block="' + (patch.block + 1) + '"]');
                            if (nextBegin) {
                                // Also include the anchor before nextBegin if present
                                var ref = nextBegin.previousSibling && nextBegin.previousSibling.tagName === 'A' ? nextBegin.previousSibling : nextBegin;
                                while (tmp.firstChild) ref.parentNode.insertBefore(tmp.firstChild, ref);
                            } else {
                                // Last resort: append to the closest page-body
                                var body = document.querySelector('.page-body');
                                while (tmp.firstChild) body.appendChild(tmp.firstChild);
                            }
                        }
                    }
                });
                // Set scroll target
                if (contentAdded) {
                    window._pendingScrollTo = '_last_page';
                    window._pendingScrollBehavior = 'instant';
                } else if (msg.scrollTo) {
                    window._pendingScrollTo = msg.scrollTo;
                }
                _clientVersion = msg.version;
                // Re-paginate + render new KaTeX/CJK
                if (typeof window._wordPaginate === 'function') window._wordPaginate();
                // Re-apply selection + marks after block-level DOM mutations
                reapplyDecorations();
            }
            es.addEventListener('update', function(e) {
                var msg = JSON.parse(e.data);
                // Track version
                if (msg.version !== undefined) _clientVersion = msg.version;
                if (msg.action === 'word-patch') {
                    // Version gap check: if we missed messages, fallback to full
                    if (msg.baseVersion !== 0 && msg.baseVersion !== _clientVersion) {
                        wordDiffUpdate(msg);
                        if (msg.version !== undefined) _clientVersion = msg.version;
                        return;
                    }
                    wordPatchUpdate(msg);
                    return;
                }
                if (msg.action === 'full') {
                    // Word: fallback diff-based update
                    if (document.querySelector('.page-wrapper[data-section]')) {
                        wordDiffUpdate(msg);
                        return;
                    }
                    // Non-Word (PPT/Excel): full body replacement
                    fetch('/').then(function(r) { return r.text(); }).then(function(html) {
                        var doc = new DOMParser().parseFromString(html, 'text/html');
                        var oldStyles = document.querySelectorAll('head style');
                        var newStyles = doc.querySelectorAll('head style');
                        oldStyles.forEach(function(s) { s.remove(); });
                        newStyles.forEach(function(s) { document.head.appendChild(s.cloneNode(true)); });
                        var scripts = document.body.querySelectorAll('script');
                        var sseScript = null;
                        scripts.forEach(function(s) { if (s.textContent.indexOf('EventSource') >= 0) sseScript = s; });
                        var targetSheetIdx = -1;
                        if (msg.scrollTo && msg.scrollTo.indexOf('data-sheet') >= 0) {
                            var m = msg.scrollTo.match(/data-sheet="(\d+)"/);
                            if (m) targetSheetIdx = parseInt(m[1]);
                        }
                        if (targetSheetIdx >= 0) {
                            doc.querySelectorAll('.sheet-content').forEach(function(s) {
                                var idx = parseInt(s.getAttribute('data-sheet'));
                                if (idx === targetSheetIdx) s.classList.add('active');
                                else s.classList.remove('active');
                            });
                            doc.querySelectorAll('.sheet-tab').forEach(function(t) {
                                var idx = parseInt(t.getAttribute('data-sheet'));
                                if (idx === targetSheetIdx) t.classList.add('active');
                                else t.classList.remove('active');
                            });
                        }
                        var savedScrollY = window.scrollY;
                        document.body.innerHTML = doc.body.innerHTML;
                        if (sseScript) document.body.appendChild(sseScript);
                        window.scrollTo(0, savedScrollY);
                        doc.body.querySelectorAll('script').forEach(function(s) {
                            if (s.textContent.indexOf('EventSource') >= 0) return;
                            var ns = document.createElement('script');
                            ns.textContent = s.textContent;
                            document.body.appendChild(ns);
                        });
                        if (msg.scrollTo && targetSheetIdx < 0) {
                            window._pendingScrollTo = msg.scrollTo;
                        }
                        // Re-apply selection + marks after the body swap destroyed previous decorations
                        reapplyDecorations();
                    });
                    return;
                }
                var slideNum = msg.slide;
                if (msg.action === 'replace') {
                    var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
                    if (el) {
                        var tmp = document.createElement('div');
                        tmp.innerHTML = msg.html;
                        var newEl = tmp.firstElementChild;
                        el.parentNode.replaceChild(newEl, el);
                        if (typeof scaleSlides === 'function') scaleSlides();
                        syncThumbs();
                        scrollToSlide(slideNum);
                    } else {
                        location.reload();
                    }
                    reapplyDecorations();
                } else if (msg.action === 'remove') {
                    var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
                    if (el) el.remove();
                    // renumber remaining slides
                    document.querySelectorAll('.slide-container').forEach(function(c, i) {
                        c.setAttribute('data-slide', i + 1);
                    });
                    syncThumbs();
                    reapplyDecorations();
                } else if (msg.action === 'add') {
                    var main = document.querySelector('.main');
                    if (main) {
                        var tmp = document.createElement('div');
                        tmp.innerHTML = msg.html;
                        var newEl = tmp.firstElementChild;
                        main.appendChild(newEl);
                        if (typeof scaleSlides === 'function') scaleSlides();
                    }
                    syncThumbs();
                    scrollToSlide(slideNum);
                    reapplyDecorations();
                }
            });
        })();
        </script>
        """;

    public WatchServer(string filePath, int port, TimeSpan? idleTimeout = null, string? initialHtml = null)
    {
        _filePath = Path.GetFullPath(filePath);
        _pipeName = GetWatchPipeName(_filePath);
        _port = port;
        _idleTimeout = idleTimeout ?? TimeSpan.FromMinutes(5);
        _tcpListener = new TcpListener(IPAddress.Loopback, _port);
        if (!string.IsNullOrEmpty(initialHtml))
            _currentHtml = initialHtml;
    }

    public static string GetWatchPipeName(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
            fullPath = fullPath.ToUpperInvariant();
        var hash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)))[..16];
        return $"officecli-watch-{hash}";
    }

    /// <summary>
    /// Check if another watch process is already running for this file.
    /// Returns the port number if running, or null if not.
    /// </summary>
    public static int? GetExistingWatchPort(string filePath)
    {
        try
        {
            int? result = null;
            var task = Task.Run(() =>
            {
                var pipeName = GetWatchPipeName(filePath);
                using var client = new System.IO.Pipes.NamedPipeClientStream(".", pipeName, System.IO.Pipes.PipeDirection.InOut);
                client.Connect(100);
                var noBom = new UTF8Encoding(false);
                using var writer = new StreamWriter(client, noBom, leaveOpen: true) { AutoFlush = true };
                writer.WriteLine("ping");
                writer.Flush();
                using var reader = new StreamReader(client, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                var response = reader.ReadLine();
                result = int.TryParse(response, out var port) ? port : 0;
            });
            return task.Wait(TimeSpan.FromSeconds(2)) ? result : null;
        }
        catch
        {
            return null; // not running
        }
    }

    public async Task RunAsync(CancellationToken externalToken = default)
    {
        // Prevent duplicate watch processes for the same file
        var existingPort = GetExistingWatchPort(_filePath);
        if (existingPort.HasValue)
        {
            var url = existingPort.Value > 0 ? $" at http://localhost:{existingPort.Value}" : "";
            throw new InvalidOperationException($"Another watch process is already running{url} for {_filePath}");
        }

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token, externalToken);
        var token = linkedCts.Token;

        _tcpListener.Start();
        Console.WriteLine($"Watch: http://localhost:{_port}");
        Console.WriteLine($"Watching: {_filePath}");
        Console.WriteLine("Press Ctrl+C to stop.");

        var pipeTask = RunPipeListenerAsync(token);
        var idleTask = RunIdleWatchdogAsync(token);

        while (!token.IsCancellationRequested)
        {
            try
            {
                var client = await _tcpListener.AcceptTcpClientAsync(token);
                _ = HandleClientAsync(client, token);
            }
            catch (OperationCanceledException) { break; }
            catch (SocketException) { break; }
            catch (ObjectDisposedException) { break; }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Watch HTTP error: {ex.Message}");
            }
        }

        // Pipe listener may not cancel promptly on Windows (WaitForConnectionAsync
        // ignores CancellationToken on some OS versions). Connect-and-drop to unblock it.
        try
        {
            using var kickPipe = new System.IO.Pipes.NamedPipeClientStream(".", _pipeName, System.IO.Pipes.PipeDirection.InOut);
            kickPipe.Connect(500);
        }
        catch { }

        try { await pipeTask; } catch (OperationCanceledException) { }
        try { await idleTask; } catch (OperationCanceledException) { }
    }

    private async Task RunIdleWatchdogAsync(CancellationToken token)
    {
        var checkInterval = TimeSpan.FromSeconds(Math.Min(30, Math.Max(1, _idleTimeout.TotalSeconds / 2)));
        while (!token.IsCancellationRequested)
        {
            await Task.Delay(checkInterval, token);
            int clientCount;
            lock (_sseLock) { clientCount = _sseClients.Count; }
            if (clientCount == 0 && DateTime.UtcNow - _lastActivityTime > _idleTimeout)
            {
                Console.WriteLine("Watch: idle timeout, shutting down.");
                _cts.Cancel();
                break;
            }
        }
    }

    private async Task RunPipeListenerAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            var server = new System.IO.Pipes.NamedPipeServerStream(
                _pipeName, System.IO.Pipes.PipeDirection.InOut,
                System.IO.Pipes.NamedPipeServerStream.MaxAllowedServerInstances,
                System.IO.Pipes.PipeTransmissionMode.Byte,
                System.IO.Pipes.PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(token);
            }
            catch (OperationCanceledException) { await server.DisposeAsync(); break; }
            catch { await server.DisposeAsync(); continue; }

            // Handle the client on a background task and immediately loop back
            // to accept another connection. This avoids a tiny window where the
            // pipe is not listening between iterations and back-to-back CLI
            // calls (e.g. multiple mark adds in a tight test loop) get refused.
            _ = Task.Run(async () =>
            {
                using (server)
                {
                    try { await HandleSinglePipeClientAsync(server, token); }
                    catch { /* ignore individual client errors */ }
                }
            }, token);
        }
    }

    private async Task HandleSinglePipeClientAsync(System.IO.Pipes.NamedPipeServerStream server, CancellationToken token)
    {
            try
            {
                var noBom = new UTF8Encoding(false);
                using var reader = new StreamReader(server, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                using var writer = new StreamWriter(server, noBom, leaveOpen: true) { AutoFlush = true };

                var message = await reader.ReadLineAsync(token);
                _lastActivityTime = DateTime.UtcNow;

                if (message == "close")
                {
                    await writer.WriteLineAsync("ok".AsMemory(), token);
                    Console.WriteLine("Watch closed by remote command.");
                    try { _tcpListener.Stop(); } catch { }
                    _cts.Cancel();
                    return;
                }
                else if (message == "ping")
                {
                    // Return port so callers can find the existing watch URL
                    await writer.WriteLineAsync(_port.ToString().AsMemory(), token);
                }
                else if (message == "get-selection")
                {
                    // Return current selection as a JSON array of paths.
                    // Empty selection → "[]". Never null.
                    string[] snapshot;
                    lock (_selectionLock) { snapshot = _currentSelection.ToArray(); }
                    var json = JsonSerializer.Serialize(snapshot, WatchSelectionJsonOptions.StringArrayInfo);
                    await writer.WriteLineAsync(json.AsMemory(), token);
                }
                else if (message == "get-marks")
                {
                    // Return {"version":N,"marks":[...]} so callers can do CAS-style
                    // detection. Empty marks → []. Never null.
                    // Uses Relaxed options so CJK content emits literal chars.
                    WatchMark[] snapshot;
                    int version;
                    lock (_marksLock)
                    {
                        snapshot = _currentMarks.ToArray();
                        version = _marksVersion;
                    }
                    var resp = new MarksResponse { Version = version, Marks = snapshot };
                    var payload = JsonSerializer.Serialize(resp, WatchMarkJsonOptions.MarksResponseInfo);
                    await writer.WriteLineAsync(payload.AsMemory(), token);
                }
                else if (message != null && message.StartsWith("mark ", StringComparison.Ordinal))
                {
                    // "mark <json>" — add a mark, return assigned id
                    var payload = message.Substring(5);
                    var resp = HandleMarkAdd(payload);
                    await writer.WriteLineAsync(resp.AsMemory(), token);
                }
                else if (message != null && message.StartsWith("unmark ", StringComparison.Ordinal))
                {
                    // "unmark <json>" — remove marks by path or all
                    var payload = message.Substring(7);
                    var resp = HandleMarkRemove(payload);
                    await writer.WriteLineAsync(resp.AsMemory(), token);
                }
                else if (message != null)
                {
                    await writer.WriteLineAsync("ok".AsMemory(), token);
                    // Try to parse as WatchMessage JSON
                    HandleWatchMessage(message);
                }
            }
            catch (OperationCanceledException) { return; }
            catch { /* ignore pipe errors */ }
    }

    private void HandleWatchMessage(string json)
    {
        try
        {
            var msg = JsonSerializer.Deserialize(json, WatchMessageJsonContext.Default.WatchMessage);
            if (msg == null) return;

            var oldHtml = _currentHtml;
            var baseVersion = _version;

            // Always update cached full HTML when provided (authoritative snapshot)
            if (!string.IsNullOrEmpty(msg.FullHtml))
            {
                _currentHtml = msg.FullHtml;
            }

            // Apply incremental patch when no full HTML was provided
            if (string.IsNullOrEmpty(msg.FullHtml))
            {
                if (msg.Action == "replace" && msg.Slide > 0 && msg.Html != null)
                    _currentHtml = PatchSlideInHtml(_currentHtml, msg.Slide, msg.Html);
                else if (msg.Action == "add" && msg.Html != null)
                    _currentHtml = AppendSlideToHtml(_currentHtml, msg.Html);
                else if (msg.Action == "remove" && msg.Slide > 0)
                    _currentHtml = RemoveSlideFromHtml(_currentHtml, msg.Slide);
            }

            _version++;

            // Reconcile all marks against the freshly updated snapshot. Flips
            // stale flags and refreshes matched_text when the underlying text
            // changed. CONSISTENCY(path-stability): same naive resolve used on
            // initial add, no fingerprint.
            ReconcileAllMarks();

            // Word: try block-level diff instead of full refresh
            if (msg.Action == "full" && !string.IsNullOrEmpty(msg.FullHtml)
                && !string.IsNullOrEmpty(oldHtml) && oldHtml.Contains("data-block=\"1\""))
            {
                var patches = ComputeWordPatches(oldHtml, msg.FullHtml);
                // Check if CSS styles changed
                var oldStyle = ExtractStyleBlock(oldHtml);
                var newStyle = ExtractStyleBlock(msg.FullHtml);
                var styleChanged = oldStyle != newStyle;

                if (patches != null || styleChanged)
                {
                    patches ??= new List<WordPatch>();
                    if (styleChanged)
                        patches.Insert(0, new WordPatch { Op = "style", Block = 0, Html = newStyle });
                    SendSseWordPatch(patches, _version, baseVersion, msg.ScrollTo);
                    return;
                }
            }

            // Forward to SSE clients (full or PPT incremental)
            SendSseEvent(msg.Action, msg.Slide, msg.Html, msg.ScrollTo, _version);
        }
        catch
        {
            // Legacy format or parse error — treat as full refresh signal
            _version++;
            SendSseEvent("full", 0, null, null, _version);
        }
    }

    // ==================== Marks ====================

    /// <summary>
    /// Add a new mark. Normalizes find: if regex flag (truthy via the find
    /// payload's "regex" field would be parsed by the CLI side; the server
    /// receives the canonical form already wrapped as r"..." or literal).
    /// However we ALSO accept the bare-find form here so that callers that
    /// don't pre-wrap still get correct behaviour. The CLI passes either
    /// the literal or a pre-wrapped r"..." string.
    /// </summary>
    internal string HandleMarkAdd(string json)
    {
        try
        {
            var req = JsonSerializer.Deserialize(json, WatchMarkJsonContext.Default.MarkRequest);
            if (req == null)
                return "{\"error\":\"invalid request\"}";

            // BUG-FUZZER-003/004: path hardening.
            //   1. Normalize: Trim() strips ASCII + Unicode whitespace from edges.
            //   2. Reject whitespace-only paths (IsNullOrWhiteSpace catches NBSP,
            //      U+3000 ideographic space, etc.).
            //   3. Require leading '/': zero-width space U+200B and BOM U+FEFF
            //      are not .NET whitespace but are never valid data-path prefixes,
            //      so a StartsWith('/') check also filters them out.
            //   4. Store the trimmed form so later `unmark --path /body/p[1]`
            //      matches what the user typed, not `" /body/p[1] "` with padding.
            // BUG-BT-R303: error messages must be actionable for AI agents — say
            // what the accepted format is, not just "invalid".
            var trimmedPath = req.Path?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(trimmedPath) || !trimmedPath.StartsWith("/"))
                return "{\"error\":\"invalid path: must start with '/' (e.g. /body/p[1] for Word, /slide[1]/shape[@id=N] for PowerPoint)\"}";

            // BUG-TESTER-002: validate color server-side. The browser sets
            // el.style.backgroundColor = mark.color verbatim, so an unsanitized
            // value injects CSS into every connected SSE client. Server is the
            // single trust boundary for both human-typed CLI and machine agents.
            // CONSISTENCY(mark-color-validation): one validator, both Add and
            // any future Set/update path must call IsValidMarkColor.
            //
            // BUG-FUZZER-001: Trim() before validation AND before storage, so
            // `"red\n"` doesn't end up stored as `"red\n"` after being accepted
            // (the validator trims for matching but used to leave the raw form
            // in the stored mark, causing a validator-vs-storage inconsistency).
            var trimmedColor = req.Color?.Trim();
            // BUG-A-R2-M01: accept bare hex (FF00FF, F0F) for consistency with the
            // rest of officecli's color parsers. The validator below requires the
            // canonical #-prefixed form, so promote 3/6/8-digit bare hex to that
            // form before validation. Anything else (named colors, rgb(...),
            // already-hashed hex) passes through unchanged.
            trimmedColor = NormalizeMarkColorInput(trimmedColor);
            // BUG-BT-R303: actionable error message — list the accepted formats
            // so AI agents can self-correct without reading the source.
            if (!string.IsNullOrEmpty(trimmedColor) && !IsValidMarkColor(trimmedColor))
                return "{\"error\":\"invalid color: accepted forms are #RGB / #RRGGBB / #RRGGBBAA hex (with or without # prefix), rgb(r,g,b), rgba(r,g,b,a), or named colors (red, blue, yellow, orange, green, purple, ...)\"}";

            var mark = new WatchMark
            {
                Path = trimmedPath,
                Find = req.Find,
                Color = string.IsNullOrEmpty(trimmedColor) ? "#ffeb3b" : trimmedColor,
                Note = req.Note,
                Tofix = req.Tofix,
                MatchedText = Array.Empty<string>(),
                Stale = false,
                CreatedAt = DateTime.UtcNow,
            };

            string assignedId;
            WatchMark[] snapshot;
            string htmlSnapshot;
            lock (_marksLock)
            {
                assignedId = _nextMarkId.ToString();
                _nextMarkId++;
                mark.Id = assignedId;
                // Snapshot _currentHtml under the lock so a concurrent
                // full-refresh can't race the resolve step.
                htmlSnapshot = _currentHtml;
                var resolved = ResolveMark(mark, htmlSnapshot);
                _currentMarks.Add(resolved);
                _marksVersion++;
                snapshot = _currentMarks.ToArray();
            }
            _lastActivityTime = DateTime.UtcNow;
            BroadcastMarkUpdate(snapshot);

            return JsonSerializer.Serialize(
                new MarkResponse { Id = assignedId },
                WatchMarkJsonContext.Default.MarkResponse);
        }
        catch
        {
            return "{\"error\":\"parse failed\"}";
        }
    }

    /// <summary>
    /// Remove marks. UnmarkRequest must have either Path set, or All=true,
    /// not both. Returns the number of marks removed.
    /// </summary>
    internal string HandleMarkRemove(string json)
    {
        try
        {
            var req = JsonSerializer.Deserialize(json, WatchMarkJsonContext.Default.UnmarkRequest);
            if (req == null) return "{\"removed\":0}";

            int removed = 0;
            WatchMark[] snapshot;
            lock (_marksLock)
            {
                if (req.All)
                {
                    removed = _currentMarks.Count;
                    _currentMarks.Clear();
                }
                else
                {
                    // BUG-FUZZER-003/004: Trim and require leading '/' for symmetry
                    // with HandleMarkAdd. Without Trim a `unmark --path " /p[1] "`
                    // would silently miss a mark added as `/p[1]` and vice versa.
                    var unmarkPath = req.Path?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(unmarkPath) && unmarkPath.StartsWith("/"))
                    {
                        removed = _currentMarks.RemoveAll(m =>
                            string.Equals(m.Path, unmarkPath, StringComparison.Ordinal));
                    }
                }
                if (removed > 0) _marksVersion++;
                snapshot = _currentMarks.ToArray();
            }
            _lastActivityTime = DateTime.UtcNow;
            if (removed > 0) BroadcastMarkUpdate(snapshot);

            return JsonSerializer.Serialize(
                new UnmarkResponse { Removed = removed },
                WatchMarkJsonContext.Default.UnmarkResponse);
        }
        catch
        {
            return "{\"removed\":0}";
        }
    }

    /// <summary>Test-only accessor for current marks snapshot.</summary>
    internal WatchMark[] GetMarksSnapshot()
    {
        lock (_marksLock) { return _currentMarks.ToArray(); }
    }

    /// <summary>Test-only accessor for the current marks version.</summary>
    internal int GetMarksVersion()
    {
        lock (_marksLock) { return _marksVersion; }
    }

    /// <summary>
    /// Test-only hook: install a full HTML snapshot synchronously and trigger
    /// mark reconciliation. Used by WatchMarkTests to verify ResolveMark without
    /// racing the pipe's "ack first, process later" ordering.
    /// </summary>
    internal void ApplyFullHtmlForTests(string html)
    {
        _currentHtml = html ?? "";
        _version++;
        ReconcileAllMarks();
    }

    // -------- Mark resolution (server-side reconcile) --------
    //
    // CONSISTENCY(path-stability): resolution uses naive positional
    // data-path lookup — no fingerprinting, no drift detection. If an
    // element is later removed or its find target no longer matches,
    // the mark is flipped to Stale=true with MatchedText=[]. Same
    // limitations as selection. grep "CONSISTENCY(path-stability)" for
    // all deferred sites that should move together if we ever switch
    // to stable IDs. See CLAUDE.md "Watch Server Rules".
    //
    // watch-isolation: this code runs pure-regex string-scraping on
    // the html snapshot already cached in _currentHtml. It does not
    // open the document, does not depend on OfficeCli.Handlers, and
    // does not reference any DOM parser. A real HTML parser would be
    // more correct but would introduce coupling; the MVP trades
    // precision for isolation and matches the browser-side
    // applyMarks() fallback behaviour.

    private static readonly System.Text.RegularExpressions.Regex _tagStripRx =
        new("<[^>]+>", System.Text.RegularExpressions.RegexOptions.Compiled);

    // BUG-TESTER-001: ResolveMark accepts arbitrary user regex via r"..." find
    // strings. A catastrophically backtracking pattern (e.g. r"(a+)+$") against
    // a long input would freeze the watch reconcile loop indefinitely. Bound
    // every user-supplied regex evaluation with this match timeout.
    private static readonly TimeSpan MarkRegexMatchTimeout = TimeSpan.FromMilliseconds(500);

    // BUG-TESTER-003: <script> and <style> bodies must be removed entirely
    // before tag-stripping, otherwise their inner text leaks into find matching
    // (e.g. find="secret" hits "<script>secret data</script>"). These regexes
    // strip the element including children, case-insensitive, dot-matches-newline.
    private static readonly System.Text.RegularExpressions.Regex _scriptBodyRx =
        new("<script\\b[^>]*>.*?</script\\s*>",
            System.Text.RegularExpressions.RegexOptions.Compiled
            | System.Text.RegularExpressions.RegexOptions.IgnoreCase
            | System.Text.RegularExpressions.RegexOptions.Singleline);
    private static readonly System.Text.RegularExpressions.Regex _styleBodyRx =
        new("<style\\b[^>]*>.*?</style\\s*>",
            System.Text.RegularExpressions.RegexOptions.Compiled
            | System.Text.RegularExpressions.RegexOptions.IgnoreCase
            | System.Text.RegularExpressions.RegexOptions.Singleline);

    // BUG-TESTER-002: server-side color whitelist for mark.color. Anything
    // accepted here gets written verbatim into el.style.backgroundColor on
    // every connected browser, so the validator must REJECT anything that
    // isn't unambiguously a color value. Three accepted shapes:
    //   1. #RGB / #RRGGBB / #RRGGBBAA hex
    //   2. rgb(r,g,b) / rgba(r,g,b,a) with numeric components
    //   3. one of the named colors in MarkNamedColors
    // CONSISTENCY(mark-color-validation): grep this tag if expanding the set.
    private static readonly System.Text.RegularExpressions.Regex _hexColorRx =
        new("^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$",
            System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _rgbFuncRx =
        new("^rgba?\\(\\s*\\d+(?:\\.\\d+)?\\s*,\\s*\\d+(?:\\.\\d+)?\\s*,\\s*\\d+(?:\\.\\d+)?(?:\\s*,\\s*\\d+(?:\\.\\d+)?)?\\s*\\)$",
            System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly HashSet<string> MarkNamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        "red", "green", "blue", "yellow", "orange", "purple", "pink", "cyan",
        "magenta", "brown", "black", "white", "gray", "grey", "lime", "teal",
        "navy", "olive", "maroon", "silver", "gold", "transparent",
    };

    // BUG-A-R2-M01 / BUG-TESTER-R302: Promote bare 3-, 6-, or 8-digit hex to
    // #-prefixed form so the validator and storage match the rest of officecli's
    // color convention. Returns the input unchanged for any other shape (named,
    // rgb(...), already #-prefixed, or null/empty). Idempotent.
    private static readonly System.Text.RegularExpressions.Regex _bareHex6Rx =
        new("^[0-9a-fA-F]{6}$", System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _bareHex3Rx =
        new("^[0-9a-fA-F]{3}$", System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _bareHex8Rx =
        new("^[0-9a-fA-F]{8}$", System.Text.RegularExpressions.RegexOptions.Compiled);
    internal static string? NormalizeMarkColorInput(string? color)
    {
        if (string.IsNullOrEmpty(color)) return color;
        if (color[0] == '#') return color;
        if (_bareHex6Rx.IsMatch(color))
            return "#" + color.ToUpperInvariant();
        if (_bareHex8Rx.IsMatch(color))
            return "#" + color.ToUpperInvariant();
        if (_bareHex3Rx.IsMatch(color))
        {
            var c = color.ToUpperInvariant();
            return $"#{c[0]}{c[0]}{c[1]}{c[1]}{c[2]}{c[2]}";
        }
        return color;
    }

    internal static bool IsValidMarkColor(string color)
    {
        if (string.IsNullOrWhiteSpace(color)) return false;
        var c = color.Trim();
        if (c.Length > 64) return false; // defensive bound
        if (MarkNamedColors.Contains(c)) return true;
        if (_hexColorRx.IsMatch(c)) return true;
        if (_rgbFuncRx.IsMatch(c)) return true;
        return false;
    }

    /// <summary>
    /// Locate the element with the given data-path in the cached HTML snapshot
    /// and return its inner HTML fragment (start tag + children + end tag).
    /// Uses bracket-depth counting of sibling tags to find the matching close.
    /// Returns null if the path is not present.
    /// </summary>
    private static string? FindDataPathInHtml(string html, string path)
    {
        if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(path)) return null;
        // Anchor the search on the data-path attribute. Path may contain [] so
        // we match it as a literal substring inside quotes.
        var marker = "data-path=\"" + path + "\"";
        var idx = html.IndexOf(marker, StringComparison.Ordinal);
        if (idx < 0) return null;
        // Walk back to the opening '<' of this element's start tag.
        var start = html.LastIndexOf('<', idx);
        if (start < 0) return null;
        // Find the end of the start tag.
        var startEnd = html.IndexOf('>', idx);
        if (startEnd < 0) return null;
        // Self-closing tag? (extremely unlikely for data-path targets but be safe)
        if (html[startEnd - 1] == '/')
            return html.Substring(start, startEnd - start + 1);
        // Extract the tag name so we can match its close.
        var tagEnd = start + 1;
        while (tagEnd < html.Length && !char.IsWhiteSpace(html[tagEnd]) && html[tagEnd] != '>')
            tagEnd++;
        var tag = html.Substring(start + 1, tagEnd - start - 1).ToLowerInvariant();
        var openToken = "<" + tag;
        var closeToken = "</" + tag;
        // Count nested open/close to find the matching close tag.
        var depth = 1;
        var cursor = startEnd + 1;
        while (cursor < html.Length && depth > 0)
        {
            var nextOpen = html.IndexOf(openToken, cursor, StringComparison.OrdinalIgnoreCase);
            var nextClose = html.IndexOf(closeToken, cursor, StringComparison.OrdinalIgnoreCase);
            if (nextClose < 0) return null;
            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                // Ensure the candidate open isn't actually part of a longer tag name
                var after = nextOpen + openToken.Length;
                if (after < html.Length && (html[after] == ' ' || html[after] == '>' || html[after] == '\t' || html[after] == '\n'))
                {
                    depth++;
                    cursor = after;
                    continue;
                }
                cursor = nextOpen + openToken.Length;
                continue;
            }
            depth--;
            cursor = nextClose + closeToken.Length;
            if (depth == 0)
            {
                // Advance past the close tag's '>'
                var gt = html.IndexOf('>', cursor);
                if (gt < 0) return null;
                return html.Substring(start, gt - start + 1);
            }
        }
        return null;
    }

    /// <summary>
    /// Extract plain text content from an HTML fragment: strip all tags, decode
    /// HTML entities, collapse whitespace minimally, and NFC-normalize. Pure
    /// regex — no DOM parser dependency.
    /// </summary>
    internal static string ExtractTextContent(string htmlFragment)
    {
        if (string.IsNullOrEmpty(htmlFragment)) return "";
        // BUG-TESTER-003: drop <script>...</script> and <style>...</style> bodies
        // BEFORE per-tag stripping. _tagStripRx only removes tags, so without
        // this step inner JS/CSS text leaks into find matching.
        var noScript = _scriptBodyRx.Replace(htmlFragment, "");
        var noStyle = _styleBodyRx.Replace(noScript, "");
        var stripped = _tagStripRx.Replace(noStyle, "");
        var decoded = System.Net.WebUtility.HtmlDecode(stripped);
        try { return decoded.Normalize(System.Text.NormalizationForm.FormC); }
        catch { return decoded; }
    }

    /// <summary>
    /// Resolve a mark against the current HTML snapshot: populate
    /// MatchedText and Stale based on whether the path still resolves
    /// and whether find still matches.
    ///
    /// Pure function: returns a new WatchMark, does not mutate the input.
    /// The caller is responsible for locking _marksLock if it's writing back
    /// into _currentMarks.
    /// </summary>
    internal static WatchMark ResolveMark(WatchMark mark, string currentHtml)
    {
        var resolved = new WatchMark
        {
            Id = mark.Id,
            Path = mark.Path,
            Find = mark.Find,
            Color = mark.Color,
            Note = mark.Note,
            Tofix = mark.Tofix,
            CreatedAt = mark.CreatedAt,
            // Defaults get overwritten below.
            MatchedText = Array.Empty<string>(),
            Stale = false,
        };

        if (string.IsNullOrEmpty(currentHtml))
        {
            // No snapshot yet (watch just started, first refresh not arrived) —
            // treat as "not resolvable yet" but don't flag stale: the CLI may
            // be adding marks before the first render. Stale stays false.
            return resolved;
        }

        var fragment = FindDataPathInHtml(currentHtml, mark.Path);
        if (fragment == null)
        {
            resolved.Stale = true;
            return resolved;
        }

        if (string.IsNullOrEmpty(mark.Find))
        {
            // Whole-element mark — no text matching needed.
            return resolved;
        }

        var text = ExtractTextContent(fragment);
        var find = mark.Find;

        // CONSISTENCY(find-regex): r"..." / r'...' raw-string prefix detection
        // matches WordHandler.Set.cs:60-61 and CommandBuilder.Mark.cs. Keep in
        // sync. grep "CONSISTENCY(find-regex)" for every project-wide site.
        bool isRegex = find.Length >= 3
            && find[0] == 'r'
            && (find[1] == '"' || find[1] == '\'')
            && find[^1] == find[1];

        if (isRegex)
        {
            var pattern = find.Substring(2, find.Length - 3);
            try
            {
                // BUG-TESTER-001: bound the match with MarkRegexMatchTimeout so a
                // catastrophic backtracker cannot freeze the reconcile loop.
                var matches = System.Text.RegularExpressions.Regex.Matches(
                    text, pattern,
                    System.Text.RegularExpressions.RegexOptions.None,
                    MarkRegexMatchTimeout);
                if (matches.Count == 0)
                {
                    resolved.Stale = true;
                    return resolved;
                }
                var list = new string[matches.Count];
                for (int i = 0; i < matches.Count; i++) list[i] = matches[i].Value;
                resolved.MatchedText = list;
                return resolved;
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
            {
                // Pattern took too long against this input → treat as stale with
                // empty matches. Future reconciles will retry against fresh HTML.
                resolved.Stale = true;
                resolved.MatchedText = Array.Empty<string>();
                return resolved;
            }
            catch
            {
                // Bad regex → treat as no match, stale.
                resolved.Stale = true;
                return resolved;
            }
        }
        else
        {
            var needle = find;
            try { needle = needle.Normalize(System.Text.NormalizationForm.FormC); } catch { }
            if (text.IndexOf(needle, StringComparison.Ordinal) < 0)
            {
                resolved.Stale = true;
                return resolved;
            }
            resolved.MatchedText = new[] { needle };
            return resolved;
        }
    }

    /// <summary>
    /// Re-run ResolveMark on every mark in the current list. Called when the
    /// cached HTML snapshot changes (document reload / full refresh). Updates
    /// each mark's MatchedText and Stale in place and bumps _marksVersion so
    /// clients that missed the change can detect it.
    /// </summary>
    private void ReconcileAllMarks()
    {
        WatchMark[] snapshot;
        lock (_marksLock)
        {
            if (_currentMarks.Count == 0) return;
            for (int i = 0; i < _currentMarks.Count; i++)
            {
                _currentMarks[i] = ResolveMark(_currentMarks[i], _currentHtml);
            }
            _marksVersion++;
            snapshot = _currentMarks.ToArray();
        }
        BroadcastMarkUpdate(snapshot);
    }

    /// <summary>Replace a single slide fragment in the full HTML by data-slide number.</summary>
    private static string PatchSlideInHtml(string html, int slideNum, string newFragment)
    {
        var (start, end) = FindSlideFragmentRange(html, slideNum);
        if (start < 0) return html;
        return string.Concat(html.AsSpan(0, start), newFragment, html.AsSpan(end));
    }

    /// <summary>Append a slide fragment before the last closing tag of the main container.</summary>
    private static string AppendSlideToHtml(string html, string fragment)
    {
        // Find the last </div> before </body> — that's the .main container's closing tag
        var bodyClose = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        if (bodyClose < 0) return html + fragment;
        // Find the </div> just before </body>
        var mainClose = html.LastIndexOf("</div>", bodyClose, StringComparison.OrdinalIgnoreCase);
        if (mainClose < 0) return html;
        return string.Concat(html.AsSpan(0, mainClose), fragment, "\n", html.AsSpan(mainClose));
    }

    /// <summary>Remove a slide fragment from the full HTML.</summary>
    private static string RemoveSlideFromHtml(string html, int slideNum)
    {
        var (start, end) = FindSlideFragmentRange(html, slideNum);
        if (start < 0) return html;
        return string.Concat(html.AsSpan(0, start), html.AsSpan(end));
    }

    /// <summary>Find the start/end character positions of a slide-container div in the HTML.</summary>
    private static (int Start, int End) FindSlideFragmentRange(string html, int slideNum)
    {
        var marker = $"data-slide=\"{slideNum}\"";
        var idx = html.IndexOf(marker, StringComparison.Ordinal);
        if (idx < 0) return (-1, -1);

        var start = html.LastIndexOf("<div ", idx, StringComparison.Ordinal);
        if (start < 0) return (-1, -1);

        // Find matching closing </div> by counting nesting
        var depth = 0;
        var pos = start;
        while (pos < html.Length)
        {
            var nextOpen = html.IndexOf("<div", pos, StringComparison.OrdinalIgnoreCase);
            var nextClose = html.IndexOf("</div>", pos, StringComparison.OrdinalIgnoreCase);

            if (nextClose < 0) break;

            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                depth++;
                pos = nextOpen + 4;
            }
            else
            {
                depth--;
                if (depth == 0)
                    return (start, nextClose + 6);
                pos = nextClose + 6;
            }
        }

        return (-1, -1);
    }

    /// <summary>Extract all &lt;style&gt; blocks from HTML head, concatenated.</summary>
    private static string? ExtractStyleBlock(string html)
    {
        var sb = new StringBuilder();
        var idx = 0;
        while (true)
        {
            var start = html.IndexOf("<style>", idx, StringComparison.OrdinalIgnoreCase);
            if (start < 0) start = html.IndexOf("<style ", idx, StringComparison.OrdinalIgnoreCase);
            if (start < 0) break;
            var end = html.IndexOf("</style>", start, StringComparison.OrdinalIgnoreCase);
            if (end < 0) break;
            end += 8; // include </style>
            sb.Append(html, start, end - start);
            idx = end;
        }
        return sb.Length > 0 ? sb.ToString() : null;
    }

    /// <summary>Split Word HTML into blocks keyed by block number. Returns dict of blockNum → content.</summary>
    private static Dictionary<int, string> SplitWordBlocks(string html)
    {
        var blocks = new Dictionary<int, string>();
        var beginRx = new System.Text.RegularExpressions.Regex(@"<span class=""wb"" data-block=""(\d+)"" style=""display:none""></span>");
        var matches = beginRx.Matches(html);
        for (int i = 0; i < matches.Count; i++)
        {
            var m = matches[i];
            var blockNum = int.Parse(m.Groups[1].Value);
            var contentStart = m.Index + m.Length;
            var endMarker = $"<span class=\"we\" data-block=\"{blockNum}\" style=\"display:none\"></span>";
            var endIdx = html.IndexOf(endMarker, contentStart, StringComparison.Ordinal);
            if (endIdx >= 0)
                blocks[blockNum] = html[contentStart..endIdx];
        }
        return blocks;
    }

    /// <summary>Compute block-level patches between old and new Word HTML. Returns null if diff is too large (fallback to full).</summary>
    internal static List<WordPatch>? ComputeWordPatches(string oldHtml, string newHtml)
    {
        // Only diff if both are Word documents with block markers
        if (string.IsNullOrEmpty(oldHtml) || string.IsNullOrEmpty(newHtml))
            return null;
        if (!oldHtml.Contains("data-block=\"1\"") || !newHtml.Contains("data-block=\"1\""))
            return null;

        var oldBlocks = SplitWordBlocks(oldHtml);
        var newBlocks = SplitWordBlocks(newHtml);

        if (oldBlocks.Count == 0 && newBlocks.Count == 0) return null;

        var patches = new List<WordPatch>();

        // Find max block number across both
        var maxBlock = 0;
        foreach (var k in oldBlocks.Keys) if (k > maxBlock) maxBlock = k;
        foreach (var k in newBlocks.Keys) if (k > maxBlock) maxBlock = k;

        for (int b = 1; b <= maxBlock; b++)
        {
            var inOld = oldBlocks.TryGetValue(b, out var oldContent);
            var inNew = newBlocks.TryGetValue(b, out var newContent);

            if (inOld && inNew)
            {
                if (oldContent != newContent)
                    patches.Add(new WordPatch { Op = "replace", Block = b, Html = newContent });
                // else: unchanged, skip
            }
            else if (!inOld && inNew)
            {
                patches.Add(new WordPatch { Op = "add", Block = b, Html = newContent });
            }
            else if (inOld && !inNew)
            {
                patches.Add(new WordPatch { Op = "remove", Block = b });
            }
        }

        if (patches.Count == 0) return null; // no changes

        // If more than 60% of blocks changed (and enough blocks to matter), fallback to full refresh
        var totalBlocks = Math.Max(oldBlocks.Count, newBlocks.Count);
        if (totalBlocks >= 5 && patches.Count > totalBlocks * 0.6)
            return null;

        return patches;
    }

    private void SendSseWordPatch(List<WordPatch> patches, int version, int baseVersion, string? scrollTo)
    {
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"word-patch\"");
        sb.Append(",\"version\":").Append(version);
        sb.Append(",\"baseVersion\":").Append(baseVersion);
        sb.Append(",\"patches\":[");
        for (int i = 0; i < patches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"op\":\"").Append(patches[i].Op).Append('"');
            sb.Append(",\"block\":").Append(patches[i].Block);
            if (patches[i].Html != null)
            {
                sb.Append(",\"html\":");
                AppendJsonString(sb, patches[i].Html!);
            }
            sb.Append('}');
        }
        sb.Append(']');
        if (scrollTo != null)
        {
            sb.Append(",\"scrollTo\":");
            AppendJsonString(sb, scrollTo);
        }
        sb.Append('}');
        BroadcastSse(sb.ToString());
    }

    private void SendSseEvent(string action, int slideNum, string? html, string? scrollTo = null, int version = 0)
    {
        // Build JSON manually to avoid dependency
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"").Append(action).Append('"');
        sb.Append(",\"slide\":").Append(slideNum);
        sb.Append(",\"version\":").Append(version);
        if (html != null)
        {
            sb.Append(",\"html\":");
            AppendJsonString(sb, html);
        }
        if (scrollTo != null)
        {
            sb.Append(",\"scrollTo\":");
            AppendJsonString(sb, scrollTo);
        }
        sb.Append('}');

        BroadcastSse(sb.ToString());
    }

    private void BroadcastSse(string sseJson)
    {
        lock (_sseLock)
        {
            var dead = new List<NetworkStream>();
            foreach (var client in _sseClients)
            {
                try
                {
                    var data = Encoding.UTF8.GetBytes($"event: update\ndata: {sseJson}\n\n");
                    client.Write(data);
                    client.Flush();
                }
                catch
                {
                    dead.Add(client);
                }
            }
            foreach (var d in dead) _sseClients.Remove(d);
        }
    }

    private static void AppendJsonString(StringBuilder sb, string value)
    {
        sb.Append('"');
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (ch < 0x20)
                        sb.Append($"\\u{(int)ch:X4}");
                    else
                        sb.Append(ch);
                    break;
            }
        }
        sb.Append('"');
    }

    private async Task HandleClientAsync(TcpClient client, CancellationToken token)
    {
        try
        {
            var stream = client.GetStream();
            var (requestLine, headers, bodyPrefix) = await ReadHttpRequestHeaderAsync(stream, token);

            if (requestLine.Contains("GET /events"))
            {
                try
                {
                    await HandleSseAsync(stream, token);
                }
                finally
                {
                    client.Close();
                }
                return;
            }

            if (requestLine.StartsWith("POST /api/selection", StringComparison.Ordinal))
            {
                await HandlePostSelectionAsync(stream, headers, bodyPrefix, token);
                client.Close();
                return;
            }

            // Default: serve current HTML (GET / and everything else)
            var html = string.IsNullOrEmpty(_currentHtml)
                ? InjectSseScript(WaitingHtml)
                : InjectSseScript(_currentHtml);
            var bodyBytes = Encoding.UTF8.GetBytes(html);
            var header = Encoding.UTF8.GetBytes(
                $"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {bodyBytes.Length}\r\nConnection: close\r\n\r\n");
            await stream.WriteAsync(header, token);
            await stream.WriteAsync(bodyBytes, token);
            client.Close();
        }
        catch
        {
            try { client.Close(); } catch { }
        }
    }

    /// <summary>
    /// Read the HTTP request line and headers, plus any body bytes that arrived in the
    /// same TCP read. Returns (requestLine, headers, bodyPrefix). Caller is responsible
    /// for reading the rest of the body using Content-Length if needed.
    /// </summary>
    private static async Task<(string requestLine, Dictionary<string, string> headers, string bodyPrefix)>
        ReadHttpRequestHeaderAsync(NetworkStream stream, CancellationToken token)
    {
        var buffer = new byte[8192];
        var sb = new StringBuilder();
        int headerEnd = -1;
        while (headerEnd < 0)
        {
            var n = await stream.ReadAsync(buffer.AsMemory(), token);
            if (n == 0) break;
            sb.Append(Encoding.UTF8.GetString(buffer, 0, n));
            headerEnd = sb.ToString().IndexOf("\r\n\r\n", StringComparison.Ordinal);
            if (sb.Length > 32 * 1024) break; // safety cap
        }

        var raw = sb.ToString();
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (headerEnd < 0)
        {
            // No header terminator — treat the whole thing as a single line
            var firstLine = raw;
            var crlf = raw.IndexOf("\r\n", StringComparison.Ordinal);
            if (crlf >= 0) firstLine = raw[..crlf];
            return (firstLine, headers, "");
        }

        var headerSection = raw[..headerEnd];
        var bodyPrefix = raw[(headerEnd + 4)..];
        var lines = headerSection.Split("\r\n");
        var requestLine = lines.Length > 0 ? lines[0] : "";
        for (int i = 1; i < lines.Length; i++)
        {
            var colon = lines[i].IndexOf(':');
            if (colon > 0)
                headers[lines[i][..colon].Trim()] = lines[i][(colon + 1)..].Trim();
        }
        return (requestLine, headers, bodyPrefix);
    }

    // Maximum size of a POST /api/selection request body. 64 KB is plenty for tens
    // of thousands of selected paths and bounds memory + read time per request.
    private const int MaxSelectionBodyBytes = 64 * 1024;
    // Hard limit on how long we'll wait for the rest of a POST body to arrive.
    // Prevents slow-loris style stalls (Content-Length advertised, body never sent).
    private static readonly TimeSpan PostBodyReadTimeout = TimeSpan.FromSeconds(3);

    private async Task HandlePostSelectionAsync(NetworkStream stream, Dictionary<string, string> headers, string bodyPrefix, CancellationToken token)
    {
        int statusCode = 204;
        string statusText = "No Content";
        string body = bodyPrefix;

        try
        {
            // Reject runaway Content-Length up front (covers FUZZER-001 slow-loris).
            int contentLength = -1;
            if (headers.TryGetValue("Content-Length", out var clStr) && int.TryParse(clStr, out var parsedCl))
            {
                if (parsedCl < 0 || parsedCl > MaxSelectionBodyBytes)
                    throw new InvalidDataException("body too large");
                contentLength = parsedCl;
            }

            // If the bodyPrefix already exceeds Content-Length, trim it. Without this,
            // an attacker could smuggle extra bytes by sending a long body in the same
            // TCP segment as the headers (FUZZER-002).
            var prefixBytes = Encoding.UTF8.GetByteCount(body);
            if (contentLength >= 0 && prefixBytes > contentLength)
            {
                var prefBytes = Encoding.UTF8.GetBytes(body);
                body = Encoding.UTF8.GetString(prefBytes, 0, contentLength);
                prefixBytes = contentLength;
            }

            // Read any missing tail bytes, bounded by both size and time.
            if (contentLength > prefixBytes)
            {
                using var readCts = CancellationTokenSource.CreateLinkedTokenSource(token);
                readCts.CancelAfter(PostBodyReadTimeout);
                var sb = new StringBuilder(body, contentLength);
                int have = prefixBytes;
                var buf = new byte[8192];
                try
                {
                    while (have < contentLength)
                    {
                        var toRead = Math.Min(buf.Length, contentLength - have);
                        var n = await stream.ReadAsync(buf.AsMemory(0, toRead), readCts.Token);
                        if (n == 0) break;
                        sb.Append(Encoding.UTF8.GetString(buf, 0, n));
                        have += n;
                        if (have > MaxSelectionBodyBytes)
                            throw new InvalidDataException("body too large");
                    }
                }
                catch (OperationCanceledException) when (!token.IsCancellationRequested)
                {
                    throw new InvalidDataException("body read timed out");
                }
                body = sb.ToString();
            }

            // Expected JSON: {"paths": ["/slide[1]/shape[2]", ...]}
            var req = JsonSerializer.Deserialize(body, WatchSelectionJsonContext.Default.SelectionRequest);
            var newSelection = req?.Paths ?? new List<string>();
            // Strip empty/null entries defensively
            newSelection = newSelection.Where(p => !string.IsNullOrEmpty(p)).ToList();

            lock (_selectionLock) { _currentSelection = newSelection; }
            _lastActivityTime = DateTime.UtcNow;

            // Broadcast to all SSE clients so other browsers can highlight in sync
            BroadcastSelectionUpdate(newSelection);
        }
        catch
        {
            statusCode = 400;
            statusText = "Bad Request";
        }

        var resp = Encoding.UTF8.GetBytes(
            $"HTTP/1.1 {statusCode} {statusText}\r\nContent-Length: 0\r\nConnection: close\r\nAccess-Control-Allow-Origin: *\r\n\r\n");
        await stream.WriteAsync(resp, token);
    }

    private void BroadcastSelectionUpdate(List<string> paths)
    {
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"selection-update\",\"paths\":[");
        for (int i = 0; i < paths.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendJsonString(sb, paths[i]);
        }
        sb.Append("]}");
        BroadcastSse(sb.ToString());
    }

    /// <summary>
    /// Wrap a WatchMark[] snapshot in a "mark-update" SSE envelope. Called
    /// after every mark add/remove, and during initial SSE client handshake.
    /// The version field is a monotonically-increasing counter that clients
    /// can use for CAS-style update detection.
    ///
    /// Uses the Relaxed encoder so CJK find/note/tofix bytes flow through
    /// as literal characters instead of \uXXXX escapes.
    /// </summary>
    private static string BuildMarkUpdateJson(WatchMark[] marks, int version)
    {
        var marksJson = JsonSerializer.Serialize(marks, WatchMarkJsonOptions.WatchMarkArrayInfo);
        return $"{{\"action\":\"mark-update\",\"version\":{version},\"marks\":{marksJson}}}";
    }

    private void BroadcastMarkUpdate(WatchMark[] marks)
    {
        int version;
        lock (_marksLock) { version = _marksVersion; }
        BroadcastSse(BuildMarkUpdateJson(marks, version));
    }

    private async Task HandleSseAsync(NetworkStream stream, CancellationToken token)
    {
        var header = Encoding.UTF8.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream; charset=utf-8\r\nCache-Control: no-cache\r\nConnection: keep-alive\r\nAccess-Control-Allow-Origin: *\r\n\r\n");
        await stream.WriteAsync(header, token);

        _lastActivityTime = DateTime.UtcNow;

        // Send the current selection immediately so the new client can highlight
        // any elements that are already selected by other browsers viewing the same
        // file. CRITICAL: this write must happen BEFORE adding the stream to
        // _sseClients. Otherwise BroadcastSse (running on another thread under
        // _sseLock) could write to the same stream at the same time we are writing
        // the initial event here, and NetworkStream is not safe for concurrent writes
        // — interleaved bytes would corrupt SSE framing.
        try
        {
            string[] snapshot;
            lock (_selectionLock) { snapshot = _currentSelection.ToArray(); }
            var sb = new StringBuilder();
            sb.Append("{\"action\":\"selection-update\",\"paths\":[");
            for (int i = 0; i < snapshot.Length; i++)
            {
                if (i > 0) sb.Append(',');
                AppendJsonString(sb, snapshot[i]);
            }
            sb.Append("]}");
            var initEvt = Encoding.UTF8.GetBytes($"event: update\ndata: {sb}\n\n");
            await stream.WriteAsync(initEvt, token);

            // Also dump the current marks snapshot so a freshly connected browser
            // immediately sees any marks the CLI has already added. Mirrors the
            // selection init dump pattern above.
            WatchMark[] markSnapshot;
            int markVersion;
            lock (_marksLock)
            {
                markSnapshot = _currentMarks.ToArray();
                markVersion = _marksVersion;
            }
            var markJson = BuildMarkUpdateJson(markSnapshot, markVersion);
            var markInitEvt = Encoding.UTF8.GetBytes($"event: update\ndata: {markJson}\n\n");
            await stream.WriteAsync(markInitEvt, token);
        }
        catch { }

        // Now safe to register: any subsequent BroadcastSse will serialize against
        // future writes via _sseLock.
        lock (_sseLock) { _sseClients.Add(stream); }

        try
        {
            while (!token.IsCancellationRequested)
            {
                await Task.Delay(30000, token);
                var heartbeat = Encoding.UTF8.GetBytes(": heartbeat\n\n");
                await stream.WriteAsync(heartbeat, token);
            }
        }
        catch { }
        finally
        {
            lock (_sseLock) { _sseClients.Remove(stream); }
        }
    }

    private static string InjectSseScript(string html)
    {
        var idx = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
            return html[..idx] + SseScript + html[idx..];
        return html + SseScript;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            _cts.Cancel();
            try { _tcpListener.Stop(); } catch { }

            // Kick the pipe listener out of WaitForConnectionAsync — it may not
            // honour CancellationToken on some Windows versions.
            try
            {
                using var kick = new System.IO.Pipes.NamedPipeClientStream(".", _pipeName, System.IO.Pipes.PipeDirection.InOut);
                kick.Connect(500);
            }
            catch { }

            // BUG-BT-003: on Unix, .NET implements named pipes as Unix domain
            // sockets at $TMPDIR/CoreFxPipe_<name>. The runtime does NOT delete
            // these on Dispose, so they accumulate in /var/folders across many
            // watch start/stop cycles (fuzzer found 302 stale files). Clean up
            // explicitly on Unix; Windows pipes are kernel objects and need no
            // file cleanup.
            if (!OperatingSystem.IsWindows())
            {
                try
                {
                    var sockPath = Path.Combine(Path.GetTempPath(), "CoreFxPipe_" + _pipeName);
                    if (File.Exists(sockPath)) File.Delete(sockPath);
                }
                catch { /* best-effort cleanup */ }
            }

            _cts.Dispose();
        }
    }
}
