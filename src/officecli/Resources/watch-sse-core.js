// watch-sse-core.js — Layer 1: Document rendering + navigation
// SSE connection, DOM updates (full/replace/add/remove), Word diff/patch,
// slide thumbnail sync, scroll management.
//
// Coupling contract with Layer 2 (watch-overlay.js):
//   - Exports window._watchEs (EventSource) for Layer 2 to listen on
//   - Calls window._watchReapplyHook() after every DOM mutation
//   - Layer 2 sets window._watchReapplyHook = reapplyDecorations

(function() {
    var es = new EventSource('/events');
    window._watchEs = es;

    var _scrollTimer = null;
    var _viewportRestoreId = 0;
    var _pendingViewportAnchor = null;
    var _manualViewportRevision = 0;

    function _noteManualViewportIntent(e) {
        if (e && e.isTrusted === false) return;
        _manualViewportRevision++;
        // A real user movement after a mutation always wins over a queued
        // restoration from that mutation.
        _pendingViewportAnchor = null;
        _viewportRestoreId++;
    }

    window.addEventListener('wheel', _noteManualViewportIntent, { passive: true, capture: true });
    window.addEventListener('touchstart', _noteManualViewportIntent, { passive: true, capture: true });
    window.addEventListener('touchmove', _noteManualViewportIntent, { passive: true, capture: true });
    window.addEventListener('pointerdown', _noteManualViewportIntent, { passive: true, capture: true });
    window.addEventListener('pointermove', function(e) {
        if (e.buttons) _noteManualViewportIntent(e);
    }, { passive: true, capture: true });
    window.addEventListener('keydown', function(e) {
        if (!e || e.isTrusted === false) return;
        if ([
            'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight',
            'PageUp', 'PageDown', 'Home', 'End', ' ',
        ].indexOf(e.key) >= 0) {
            _noteManualViewportIntent(e);
        }
    }, true);

    function _isScrollable(el) {
        if (!el) return false;
        var style = getComputedStyle(el);
        return /(auto|scroll|overlay)/.test(style.overflowY + ' ' + style.overflowX)
            && (el.scrollHeight > el.clientHeight || el.scrollWidth > el.clientWidth);
    }

    function _activeViewportScroller() {
        var pptMain = document.querySelector('.main');
        if (_isScrollable(pptMain)) return pptMain;
        var activeSheet = document.querySelector('.sheet-content.active');
        var tableWrapper = activeSheet && activeSheet.querySelector('.table-wrapper');
        if (_isScrollable(tableWrapper)) return tableWrapper;
        return document.scrollingElement || document.documentElement;
    }

    function _viewportRect(scroller) {
        if (scroller === document.scrollingElement
            || scroller === document.documentElement
            || scroller === document.body) {
            return { top: 0, left: 0, right: innerWidth, bottom: innerHeight };
        }
        return scroller.getBoundingClientRect();
    }

    function _scrollerPosition(scroller) {
        if (scroller === document.scrollingElement
            || scroller === document.documentElement
            || scroller === document.body) {
            return { top: window.scrollY, left: window.scrollX };
        }
        return { top: scroller.scrollTop, left: scroller.scrollLeft };
    }

    function _setScrollerPosition(scroller, top, left) {
        if (scroller === document.scrollingElement
            || scroller === document.documentElement
            || scroller === document.body) {
            window.scrollTo(left, top);
            return;
        }
        scroller.scrollTop = top;
        scroller.scrollLeft = left;
    }

    function _visiblePathElements(scroller, viewport) {
        var scope = scroller === document.scrollingElement
            || scroller === document.documentElement
            || scroller === document.body
            ? document
            : scroller;
        return Array.prototype.filter.call(
            scope.querySelectorAll('[data-path]'),
            function(el) {
                if (el.closest('.sidebar,.thumb,.sheet-tabs')) return false;
                var rect = el.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0
                    && rect.bottom > viewport.top && rect.top < viewport.bottom
                    && rect.right > viewport.left && rect.left < viewport.right;
            });
    }

    function _captureViewportAnchor() {
        var scroller = _activeViewportScroller();
        if (!scroller) return null;
        var viewport = _viewportRect(scroller);
        var centerY = (viewport.top + viewport.bottom) / 2;
        var centerX = (viewport.left + viewport.right) / 2;
        var candidates = _visiblePathElements(scroller, viewport);
        var best = null;
        var bestScore = Infinity;
        for (var i = 0; i < candidates.length; i++) {
            var rect = candidates[i].getBoundingClientRect();
            var score = Math.abs((rect.top + rect.bottom) / 2 - centerY)
                + Math.abs((rect.left + rect.right) / 2 - centerX) * 0.2;
            // Prefer the more specific nested data-path when two candidates
            // occupy the same visual region.
            score -= (candidates[i].getAttribute('data-path') || '').split('/').length * 0.01;
            if (score < bestScore) {
                best = candidates[i];
                bestScore = score;
            }
        }
        var position = _scrollerPosition(scroller);
        var rect = best && best.getBoundingClientRect();
        var slide = best && best.closest('.slide-container');
        var slideRect = slide && slide.getBoundingClientRect();
        var sheet = document.querySelector('.sheet-content.active');
        var siblings = candidates;
        var bestIndex = best ? siblings.indexOf(best) : -1;
        return {
            element: best,
            path: best && best.getAttribute('data-path'),
            before: bestIndex > 0 ? siblings[bestIndex - 1] : null,
            beforePath: bestIndex > 0
                ? siblings[bestIndex - 1].getAttribute('data-path') : null,
            after: bestIndex >= 0 && bestIndex + 1 < siblings.length
                ? siblings[bestIndex + 1] : null,
            afterPath: bestIndex >= 0 && bestIndex + 1 < siblings.length
                ? siblings[bestIndex + 1].getAttribute('data-path') : null,
            offsetTop: rect ? rect.top - viewport.top : null,
            offsetLeft: rect ? rect.left - viewport.left : null,
            scrollTop: position.top,
            scrollLeft: position.left,
            slideElement: slide,
            slideNumber: slide ? parseInt(slide.getAttribute('data-slide')) || 0 : 0,
            slideOffsetTop: slideRect ? slideRect.top - viewport.top : null,
            sheetIndex: sheet ? parseInt(sheet.getAttribute('data-sheet')) : null,
            manualRevision: _manualViewportRevision,
        };
    }

    function _preferScrollTarget(candidates) {
        if (!candidates.length) return null;
        // PowerPoint thumbnails clone marked/selected slide content. Prefer the
        // real slide in .main so explicit goto/focus never navigates the
        // sidebar clone. Excel similarly prefers the currently active sheet.
        for (var i = 0; i < candidates.length; i++) {
            if (candidates[i].closest('.main')
                && !candidates[i].closest('.sidebar,.thumb')) {
                return candidates[i];
            }
        }
        for (var i = 0; i < candidates.length; i++) {
            if (candidates[i].closest('.sheet-content.active')) {
                return candidates[i];
            }
        }
        for (var i = 0; i < candidates.length; i++) {
            if (!candidates[i].closest('.sidebar,.thumb,.sheet-tabs')) {
                return candidates[i];
            }
        }
        return candidates[0];
    }

    function _findPathElement(path) {
        if (!path) return null;
        var elements = document.querySelectorAll('[data-path]');
        var candidates = [];
        for (var i = 0; i < elements.length; i++) {
            if (elements[i].getAttribute('data-path') === path) {
                candidates.push(elements[i]);
            }
        }
        if (!candidates.length) {
            var parentPath = path;
            while (parentPath.lastIndexOf('/') > 0 && !candidates.length) {
                parentPath = parentPath.substring(0, parentPath.lastIndexOf('/'));
                for (var j = 0; j < elements.length; j++) {
                    if (elements[j].getAttribute('data-path') === parentPath) {
                        candidates.push(elements[j]);
                    }
                }
            }
        }
        return _preferScrollTarget(candidates);
    }

    function _queryScrollTarget(selector) {
        var candidates = Array.prototype.slice.call(
            document.querySelectorAll(selector));
        return _preferScrollTarget(candidates);
    }

    function _scrollExplicitTarget(target) {
        if (!target) return;
        var sheet = target.closest('.sheet-content');
        if (sheet && !sheet.classList.contains('active')) {
            var sheetIndex = parseInt(sheet.getAttribute('data-sheet'));
            document.querySelectorAll('.sheet-content').forEach(function(item) {
                item.classList.toggle(
                    'active',
                    parseInt(item.getAttribute('data-sheet')) === sheetIndex);
            });
            document.querySelectorAll('.sheet-tab').forEach(function(item) {
                item.classList.toggle(
                    'active',
                    parseInt(item.getAttribute('data-sheet')) === sheetIndex);
            });
        }
        target.scrollIntoView({
            behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches
                ? 'auto' : 'smooth',
            block: 'center',
            inline: 'center',
        });
    }

    function _closestSurvivingSlide(num) {
        var slides = Array.prototype.slice.call(
            document.querySelectorAll('.main > .slide-container'));
        if (!slides.length) return null;
        var best = slides[0], bestDistance = Infinity;
        for (var i = 0; i < slides.length; i++) {
            var current = parseInt(slides[i].getAttribute('data-slide')) || (i + 1);
            var distance = Math.abs(current - num);
            if (distance < bestDistance) {
                best = slides[i];
                bestDistance = distance;
            }
        }
        return best;
    }

    function _restoreViewportAnchor(anchor) {
        if (!anchor || anchor.manualRevision !== _manualViewportRevision) return;

        if (Number.isInteger(anchor.sheetIndex)) {
            document.querySelectorAll('.sheet-content').forEach(function(sheet) {
                sheet.classList.toggle(
                    'active',
                    parseInt(sheet.getAttribute('data-sheet')) === anchor.sheetIndex);
            });
            document.querySelectorAll('.sheet-tab').forEach(function(tab) {
                tab.classList.toggle(
                    'active',
                    parseInt(tab.getAttribute('data-sheet')) === anchor.sheetIndex);
            });
        }

        var scroller = _activeViewportScroller();
        if (!scroller) return;
        var viewport = _viewportRect(scroller);
        var target = anchor.element && anchor.element.isConnected
            ? anchor.element
            : _findPathElement(anchor.path);
        if (!target) {
            target = anchor.before && anchor.before.isConnected
                ? anchor.before : _findPathElement(anchor.beforePath);
        }
        if (!target) {
            target = anchor.after && anchor.after.isConnected
                ? anchor.after : _findPathElement(anchor.afterPath);
        }

        if (target && anchor.offsetTop !== null && anchor.offsetLeft !== null) {
            var targetRect = target.getBoundingClientRect();
            var current = _scrollerPosition(scroller);
            _setScrollerPosition(
                scroller,
                current.top + (targetRect.top - viewport.top - anchor.offsetTop),
                current.left + (targetRect.left - viewport.left - anchor.offsetLeft));
            return;
        }

        var slide = anchor.slideElement && anchor.slideElement.isConnected
            ? anchor.slideElement
            : _closestSurvivingSlide(anchor.slideNumber);
        if (slide && anchor.slideOffsetTop !== null) {
            var slideRect = slide.getBoundingClientRect();
            var current = _scrollerPosition(scroller);
            _setScrollerPosition(
                scroller,
                current.top + (slideRect.top - viewport.top - anchor.slideOffsetTop),
                anchor.scrollLeft);
            return;
        }

        _setScrollerPosition(scroller, anchor.scrollTop, anchor.scrollLeft);
    }

    function _cancelViewportRestore() {
        _pendingViewportAnchor = null;
        _viewportRestoreId++;
    }

    function _queueViewportRestore(anchor, rendererWillSignal) {
        if (!anchor) return;
        var id = ++_viewportRestoreId;
        _pendingViewportAnchor = anchor;
        if (rendererWillSignal) return;
        requestAnimationFrame(function() {
            if (id !== _viewportRestoreId || _pendingViewportAnchor !== anchor) return;
            _pendingViewportAnchor = null;
            _restoreViewportAnchor(anchor);
        });
    }

    window._watchRestorePendingViewport = function() {
        var anchor = _pendingViewportAnchor;
        if (!anchor) return;
        var id = _viewportRestoreId;
        requestAnimationFrame(function() {
            if (id !== _viewportRestoreId || _pendingViewportAnchor !== anchor) return;
            _pendingViewportAnchor = null;
            _restoreViewportAnchor(anchor);
        });
    };

    function _callReapplyHook() {
        if (typeof window._watchReapplyHook === 'function') window._watchReapplyHook();
    }

    // innerHTML does not execute <script> tags, and re-creating scripts without
    // preserving the type attribute breaks ES modules (e.g. model3d / three.js).
    // Walks the subtree, replaces each <script> with a fresh element that copies
    // every attribute + textContent (or src) so the browser actually runs it.
    function _executeScripts(root) {
        if (!root) return;
        var scripts = root.querySelectorAll ? root.querySelectorAll('script') : [];
        for (var i = 0; i < scripts.length; i++) {
            var s = scripts[i];
            var ns = document.createElement('script');
            for (var j = 0; j < s.attributes.length; j++) {
                var a = s.attributes[j];
                ns.setAttribute(a.name, a.value);
            }
            if (s.src) ns.src = s.src;
            else ns.textContent = s.textContent;
            s.parentNode.replaceChild(ns, s);
        }
    }

    function _replaceDocumentBody(msg) {
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
            // Preserve current active sheet if no explicit target
            if (targetSheetIdx < 0) {
                var curActive = document.querySelector('.sheet-tab.active');
                if (curActive) targetSheetIdx = parseInt(curActive.getAttribute('data-sheet')) || 0;
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
            // Capture immediately before the DOM swap so a manual scroll made
            // while fetch('/') was in flight wins over older state.
            var viewportAnchor = _captureViewportAnchor();
            document.body.innerHTML = doc.body.innerHTML;
            if (sseScript) document.body.appendChild(sseScript);
            doc.body.querySelectorAll('script').forEach(function(s) {
                if (s.textContent.indexOf('EventSource') >= 0) return;
                var ns = document.createElement('script');
                for (var j = 0; j < s.attributes.length; j++) {
                    var a = s.attributes[j];
                    ns.setAttribute(a.name, a.value);
                }
                if (s.src) ns.src = s.src;
                else ns.textContent = s.textContent;
                document.body.appendChild(ns);
            });
            // Re-apply selection + marks after the body swap
            _callReapplyHook();
            _queueViewportRestore(viewportAnchor, false);
        });
    }

    function scrollToSlide(num) {
        _cancelViewportRestore();
        if (_scrollTimer !== null) cancelAnimationFrame(_scrollTimer);
        _scrollTimer = requestAnimationFrame(function() {
            _scrollTimer = null;
            var target = document.querySelector('.slide-container[data-slide="' + num + '"]');
            if (target) target.scrollIntoView({
                behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches
                    ? 'auto' : 'smooth',
                block: 'center',
            });
        });
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
        fetch('/').then(function(r) { return r.text(); }).then(function(html) {
            var viewportAnchor = _captureViewportAnchor();
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
            // Normal document updates preserve the latest semantic viewport
            // anchor. Explicit navigation arrives as action="scroll", never as
            // a mutation's legacy scrollTo hint.
            window._pendingScrollTo = null;
            window._pendingScrollBehavior = null;
            _queueViewportRestore(viewportAnchor, true);
            // Re-paginate (will also re-scale and remove freeze)
            if (typeof window._wordPaginate === 'function') window._wordPaginate();
            else {
                var f=document.getElementById('_sse_freeze');
                if(f)f.remove();
                window._watchRestorePendingViewport();
            }
            // Re-apply selection + marks after DOM swap
            _callReapplyHook();
        });
    }

    // Track version for gap detection
    var _clientVersion = 0;

    // Apply server-side block patches directly to DOM
    function wordPatchUpdate(msg) {
        var viewportAnchor = _captureViewportAnchor();
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
        window._pendingScrollTo = null;
        window._pendingScrollBehavior = null;
        _queueViewportRestore(viewportAnchor, true);
        _clientVersion = msg.version;
        // Re-paginate + render new KaTeX/CJK
        if (typeof window._wordPaginate === 'function') window._wordPaginate();
        else {
            var f=document.getElementById('_sse_freeze');
            if(f)f.remove();
            window._watchRestorePendingViewport();
        }
        // Re-apply selection + marks after block-level DOM mutations
        _callReapplyHook();
    }

    // The server switched to a different document in place (POST /api/switch).
    // The current DOM, per-format scripts (Word pagination etc.) and styles
    // all belong to the old document, so a full page reload is the only
    // correct reaction here — GET / re-serves the new document's HTML.
    // Embedders (which strip or replace this script) handle doc-switched
    // with their own state machine instead.
    es.addEventListener('doc-switched', function() {
        location.reload();
    });

    // Main SSE listener for DOM-swap events
    es.addEventListener('update', function(e) {
        var msg = JSON.parse(e.data);
        // Scroll-only: navigate the viewer without mutating DOM/styles.
        // Sent by the `goto` command. Word path is selector-based; for
        // PPT use scrollToSlide if scrollTo matches /slide\[N\]/.
        if (msg.action === 'scroll' && (msg.scrollTo || msg.scrollPath)) {
            _cancelViewportRestore();
            if (msg.scrollPath) {
                _scrollExplicitTarget(_findPathElement(msg.scrollPath));
                return;
            }
            var sel = msg.scrollTo;
            var slideMatch = sel.match(/data-slide="(\d+)"/);
            if (slideMatch) { scrollToSlide(parseInt(slideMatch[1])); return; }
            try {
                _scrollExplicitTarget(_queryScrollTarget(sel));
            } catch (e) { /* invalid selector — silent */ }
            return;
        }
        // Track version — save prevVersion BEFORE updating so gap checks
        // compare against the version we actually have, not the incoming one.
        var prevVersion = _clientVersion;
        if (msg.version !== undefined) _clientVersion = msg.version;
        if (msg.action === 'word-patch') {
            // Version gap check: if we missed messages, fallback to full
            // Skip when prevVersion===0 (fresh client — no messages seen yet)
            if (prevVersion > 0 && msg.baseVersion !== 0 && msg.baseVersion !== prevVersion) {
                wordDiffUpdate(msg);
                return;
            }
            wordPatchUpdate(msg);
            return;
        }
        if (msg.action === 'excel-patch') {
            var viewportAnchor = _captureViewportAnchor();
            // Version gap check: if we missed messages, fallback to full reload
            // Skip when prevVersion===0 (fresh client — no messages seen yet)
            if (prevVersion > 0 && msg.baseVersion !== 0 && msg.baseVersion !== prevVersion) {
                location.reload();
                return;
            }
            // Apply style patch if present
            msg.patches.forEach(function(patch) {
                if (patch.op === 'style') {
                    var oldStyles = document.querySelectorAll('head style');
                    oldStyles.forEach(function(s) { s.remove(); });
                    var tmp = document.createElement('div');
                    tmp.innerHTML = patch.html;
                    var styles = tmp.querySelectorAll('style');
                    styles.forEach(function(s) { document.head.appendChild(s.cloneNode(true)); });
                    return;
                }
                var existing = document.querySelector('tr[data-row="' + patch.row + '"]');
                if (patch.op === 'replace' && existing) {
                    var tmp = document.createElement('tbody');
                    tmp.innerHTML = patch.html;
                    var newRow = tmp.firstElementChild;
                    if (newRow) existing.parentNode.replaceChild(newRow, existing);
                } else if (patch.op === 'remove' && existing) {
                    existing.remove();
                } else if (patch.op === 'add' && !existing) {
                    // Find the tbody in the correct sheet and insert at sorted position
                    var parts = patch.row.split('-');
                    var sheetDiv = document.querySelector('.sheet-content[data-sheet="' + parts[0] + '"]');
                    if (sheetDiv) {
                        var tbody = sheetDiv.querySelector('tbody');
                        if (tbody) {
                            var tmp = document.createElement('tbody');
                            tmp.innerHTML = patch.html;
                            var newRow = tmp.firstElementChild;
                            if (newRow) {
                                // Insert before the first row with a higher row number
                                var newNum = parseInt(parts[1]);
                                var inserted = false;
                                var rows = tbody.querySelectorAll('tr[data-row]');
                                for (var ri = 0; ri < rows.length; ri++) {
                                    var rp = rows[ri].getAttribute('data-row').split('-');
                                    if (parseInt(rp[1]) > newNum) {
                                        tbody.insertBefore(newRow, rows[ri]);
                                        inserted = true;
                                        break;
                                    }
                                }
                                if (!inserted) tbody.appendChild(newRow);
                            }
                        }
                    }
                }
            });
            if (msg.version !== undefined) _clientVersion = msg.version;
            _callReapplyHook();
            _queueViewportRestore(viewportAnchor, false);
            return;
        }
        if (msg.action === 'full') {
            // Word: fallback diff-based update
            if (document.querySelector('.page-wrapper[data-section]')) {
                wordDiffUpdate(msg);
                return;
            }
            // Defer full body replacement while a drag is in progress
            if (window._isDragging) {
                var _deferredMsg = msg;
                function _applyWhenIdle() {
                    if (window._isDragging) { setTimeout(_applyWhenIdle, 100); return; }
                    es.dispatchEvent(new MessageEvent('update', { data: JSON.stringify(_deferredMsg) }));
                }
                setTimeout(_applyWhenIdle, 100);
                return;
            }
            // Non-Word (PPT/Excel): full body replacement
            _replaceDocumentBody(msg);
            return;
        }
        var slideNum = msg.slide;
        if (msg.action === 'replace') {
            var viewportAnchor = _captureViewportAnchor();
            var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
            if (el) {
                var tmp = document.createElement('div');
                tmp.innerHTML = msg.html;
                var newEl = tmp.firstElementChild;
                el.parentNode.replaceChild(newEl, el);
                _executeScripts(newEl);
                if (typeof scaleSlides === 'function') scaleSlides();
                syncThumbs();
            } else {
                location.reload();
            }
            _callReapplyHook();
            _queueViewportRestore(viewportAnchor, false);
        } else if (msg.action === 'remove') {
            var viewportAnchor = _captureViewportAnchor();
            var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
            if (el) el.remove();
            // renumber remaining slides
            document.querySelectorAll('.slide-container').forEach(function(c, i) {
                c.setAttribute('data-slide', i + 1);
            });
            syncThumbs();
            _callReapplyHook();
            _queueViewportRestore(viewportAnchor, false);
        } else if (msg.action === 'add') {
            var viewportAnchor = _captureViewportAnchor();
            var main = document.querySelector('.main');
            if (main) {
                var tmp = document.createElement('div');
                tmp.innerHTML = msg.html;
                var newEl = tmp.firstElementChild;
                main.appendChild(newEl);
                _executeScripts(newEl);
                if (typeof scaleSlides === 'function') scaleSlides();
            }
            syncThumbs();
            _callReapplyHook();
            _queueViewportRestore(viewportAnchor, false);
        }
    });
})();
