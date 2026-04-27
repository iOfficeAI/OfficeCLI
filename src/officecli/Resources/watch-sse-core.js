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

    function _callReapplyHook() {
        if (typeof window._watchReapplyHook === 'function') window._watchReapplyHook();
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
            // Re-apply selection + marks after the body swap
            _callReapplyHook();
        });
    }

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
            // Re-apply selection + marks after DOM swap
            _callReapplyHook();
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
        _callReapplyHook();
    }

    // Main SSE listener for DOM-swap events
    es.addEventListener('update', function(e) {
        var msg = JSON.parse(e.data);
        // Scroll-only: navigate the viewer without mutating DOM/styles.
        // Sent by the `goto` command. Word path is selector-based; for
        // PPT use scrollToSlide if scrollTo matches /slide\[N\]/.
        if (msg.action === 'scroll' && msg.scrollTo) {
            var sel = msg.scrollTo;
            var slideMatch = sel.match(/data-slide="(\d+)"/);
            if (slideMatch) { scrollToSlide(parseInt(slideMatch[1])); return; }
            try {
                var t = document.querySelector(sel);
                if (t) t.scrollIntoView({ behavior: 'smooth', block: 'center' });
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
            if (msg.scrollTo) window._pendingScrollTo = msg.scrollTo;
            if (msg.version !== undefined) _clientVersion = msg.version;
            _callReapplyHook();
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
            _callReapplyHook();
        } else if (msg.action === 'remove') {
            var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
            if (el) el.remove();
            // renumber remaining slides
            document.querySelectorAll('.slide-container').forEach(function(c, i) {
                c.setAttribute('data-slide', i + 1);
            });
            syncThumbs();
            _callReapplyHook();
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
            _callReapplyHook();
        }
    });
})();
