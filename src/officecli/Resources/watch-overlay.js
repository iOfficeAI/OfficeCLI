// watch-overlay.js — Layer 2: Overlay / decoration layer
// Selection highlighting, marks (find/regex), rubber-band box selection,
// CSS injection, and the reapply hook.
//
// Depends on Layer 1 (watch-sse-core.js) exporting:
//   - window._watchEs (EventSource) — used to listen for selection-update / mark-update
// Registers:
//   - window._watchReapplyHook — called by Layer 1 after every DOM mutation
//
// Future additions: revision panel, lightweight editing (drag, text edit)

(function() {
    var es = window._watchEs;

    // ===== Selection sync =====
    // Single source of truth: server's currentSelection. We keep a local
    // mirror updated by the server's SSE 'selection-update' broadcasts so
    // that we can re-apply highlights after every DOM swap.
    var _selection = [];

    // Detect if selected cell paths form a contiguous rectangle.
    // Returns {sheet, minC, maxC, minR, maxR, cells} or null.
    function _detectRect(paths) {
        if (paths.length === 0) return null;
        var sheet = null, minC = Infinity, maxC = -Infinity, minR = Infinity, maxR = -Infinity;
        var cells = [];
        for (var i = 0; i < paths.length; i++) {
            var c = _parseCellPath(paths[i]);
            if (!c) return null;
            if (!sheet) sheet = c.sheet; else if (c.sheet !== sheet) return null;
            var cn = _colToNum(c.col);
            if (cn < minC) minC = cn; if (cn > maxC) maxC = cn;
            if (c.row < minR) minR = c.row; if (c.row > maxR) maxR = c.row;
            cells.push({ col: cn, row: c.row, path: paths[i] });
        }
        if (cells.length !== (maxC - minC + 1) * (maxR - minR + 1)) return null;
        if (cells.length < 2) return null; // single cell uses individual styling
        return { sheet: sheet, minC: minC, maxC: maxC, minR: minR, maxR: maxR, cells: cells };
    }

    var _SEL_CLASSES = ['officecli-selected', 'officecli-sel-range', 'officecli-sel-rowcol', 'officecli-sel-handle'];

    function applySelectionToDom() {
        // Clear all selection classes + inline box-shadow from previous range
        var allSel = _SEL_CLASSES.map(function(c) { return '.' + c; }).join(',');
        document.querySelectorAll(allSel).forEach(function(el) {
            _SEL_CLASSES.forEach(function(c) { el.classList.remove(c); });
            el.style.boxShadow = '';
        });
        if (_selection.length === 0) return;

        // Try rectangular range styling (Excel-native look)
        var rect = _detectRect(_selection);
        if (rect) {
            // Highlight row/col headers (crosshair for entire range)
            for (var r = rect.minR; r <= rect.maxR; r++) {
                try {
                    var rs = '[data-path="' + (rect.sheet + '/row[' + r + ']').replace(/"/g, '\\"') + '"]';
                    document.querySelectorAll(rs).forEach(function(th) { th.classList.add('officecli-selected'); });
                } catch(e) {}
            }
            for (var c = rect.minC; c <= rect.maxC; c++) {
                try {
                    var cs = '[data-path="' + (rect.sheet + '/col[' + _numToCol(c) + ']').replace(/"/g, '\\"') + '"]';
                    document.querySelectorAll(cs).forEach(function(th) { th.classList.add('officecli-selected'); });
                } catch(e) {}
            }
            // Apply range fill + inset box-shadow for perimeter (no layout shift)
            var B = '#217346', W = 2; // border color and width
            for (var i = 0; i < rect.cells.length; i++) {
                var cell = rect.cells[i];
                try {
                    var sel = '[data-path="' + cell.path.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"]';
                    document.querySelectorAll(sel).forEach(function(el) {
                        el.classList.add('officecli-sel-range');
                        // Build inset box-shadow for edge borders
                        var shadows = [];
                        if (cell.row === rect.minR) shadows.push('inset 0 '+W+'px 0 '+B);
                        if (cell.row === rect.maxR) shadows.push('inset 0 -'+W+'px 0 '+B);
                        if (cell.col === rect.minC) shadows.push('inset '+W+'px 0 0 '+B);
                        if (cell.col === rect.maxC) shadows.push('inset -'+W+'px 0 0 '+B);
                        if (shadows.length > 0) el.style.boxShadow = shadows.join(',');
                        if (cell.row === rect.maxR && cell.col === rect.maxC)
                            el.classList.add('officecli-sel-handle');
                    });
                } catch(e) {}
            }
            return;
        }

        // Fallback: individual cell styling (non-contiguous / mixed paths)
        _selection.forEach(function(path) {
            try {
                var sel = '[data-path="' + path.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"]';
                document.querySelectorAll(sel).forEach(function(el) {
                    el.classList.add('officecli-selected');
                    // Row header: highlight row cells with stronger fill
                    var rowMatch = path.match(/^(\/[^/]+)\/row\[(\d+)\]$/);
                    if (rowMatch && el.tagName === 'TH') {
                        var tr = el.closest('tr');
                        if (tr) tr.querySelectorAll('td[data-path]').forEach(function(td) {
                            td.classList.add('officecli-sel-rowcol');
                        });
                    }
                    // Col header: highlight column cells with stronger fill
                    var colMatch = path.match(/^(\/[^/]+)\/col\[([A-Za-z]+)\]$/);
                    if (colMatch && el.tagName === 'TH') {
                        var sheet = colMatch[1], col = colMatch[2];
                        var re = new RegExp('^' + sheet.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\/' + col + '\\d+$', 'i');
                        document.querySelectorAll('td[data-path]').forEach(function(td) {
                            if (re.test(td.getAttribute('data-path')))
                                td.classList.add('officecli-sel-rowcol');
                        });
                    }
                    // Cell: crosshair headers
                    var cellMatch = _parseCellPath(path);
                    if (cellMatch && el.tagName === 'TD') {
                        try {
                            var rSel = '[data-path="' + (cellMatch.sheet + '/row[' + cellMatch.row + ']').replace(/"/g, '\\"') + '"]';
                            document.querySelectorAll(rSel).forEach(function(th) { th.classList.add('officecli-selected'); });
                            var cSel = '[data-path="' + (cellMatch.sheet + '/col[' + cellMatch.col + ']').replace(/"/g, '\\"') + '"]';
                            document.querySelectorAll(cSel).forEach(function(th) { th.classList.add('officecli-selected'); });
                        } catch(e2) {}
                    }
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

    // ===== Excel cell range helpers =====
    var _anchor = null; // {sheet, col, row} — anchor for shift-range and drag
    var _cellDrag = null; // active cell-to-cell drag state
    var _headerDrag = null; // active row/col header drag state

    function _parseCellPath(path) {
        var m = path.match(/^(\/[^/]+)\/([A-Za-z]+)(\d+)$/);
        if (!m) return null;
        var row = parseInt(m[3], 10);
        if (row < 1) return null;
        return { sheet: m[1], col: m[2].toUpperCase(), row: row };
    }
    function _colToNum(col) {
        var n = 0;
        for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
        return n;
    }
    function _numToCol(num) {
        var s = '';
        while (num > 0) { var r = (num - 1) % 26; s = String.fromCharCode(65 + r) + s; num = Math.floor((num - 1) / 26); }
        return s;
    }
    function _expandCellRange(sheet, col1, row1, col2, row2) {
        var minC = Math.min(_colToNum(col1), _colToNum(col2));
        var maxC = Math.max(_colToNum(col1), _colToNum(col2));
        var minR = Math.min(row1, row2), maxR = Math.max(row1, row2);
        var paths = [];
        for (var r = minR; r <= maxR; r++)
            for (var c = minC; c <= maxC; c++)
                paths.push(sheet + '/' + _numToCol(c) + r);
        return paths;
    }
    // Deduplicate paths while preserving order
    function _uniquePaths(arr) {
        var seen = {}, out = [];
        for (var i = 0; i < arr.length; i++) {
            if (!seen[arr[i]]) { seen[arr[i]] = true; out.push(arr[i]); }
        }
        return out;
    }

    // Inject selection + mark highlight CSS
    (function() {
        var style = document.createElement('style');
        style.textContent =
            // Range fill: light gray like real Excel (box-shadow for borders, no layout shift)
            'td.officecli-sel-range{' +
                'background:rgba(33,115,70,0.12) !important;' +
                'position:relative;' +
            '}' +
            // Row/col header selection: stronger fill for entire row/column
            'td.officecli-sel-rowcol{' +
                'background:rgba(33,115,70,0.20) !important;' +
            '}' +
            // Fill handle: small square at bottom-right corner of range
            'td.officecli-sel-handle::after{' +
                'content:"";position:absolute;right:-4px;bottom:-4px;' +
                'width:7px;height:7px;background:#217346;' +
                'border:1px solid #fff;z-index:1001;' +
            '}' +
            // Individual cell selection (non-contiguous / Ctrl+click fallback)
            'td.officecli-selected{' +
                'outline:2px solid #217346 !important;' +
                'outline-offset:-2px;' +
                'position:relative;' +
                'z-index:1000;' +
            '}' +
            'td.officecli-selected::after{' +
                'content:"";position:absolute;right:-4px;bottom:-4px;' +
                'width:7px;height:7px;background:#217346;' +
                'border:1px solid #fff;z-index:1001;' +
            '}' +
            // Header crosshair: dark green background like real Excel
            'th.officecli-selected{' +
                'background:#217346 !important;color:#fff !important;' +
            '}' +
            // Non-cell fallback (pptx/docx shapes)
            ':not(td):not(th).officecli-selected{' +
                'outline:2px solid #2196f3 !important;' +
                'outline-offset:2px;' +
                'box-shadow:0 0 12px rgba(33,150,243,0.6) !important;' +
                'z-index:1000;' +
            '}' +
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
            // still applies visual styling to the content.
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
        // Scope mark lookup to the main slide container only. The sidebar
        // thumbs are JS-cloned from .main and end up sharing the same
        // [data-path] values; document.querySelector would otherwise
        // hit the thumb (DOM-order first) and the real preview would
        // never receive the mark. See R4 trial bug.
        var _markRoot = document.querySelector('.main') || document;
        for (var mi = 0; mi < _marks.length; mi++) {
            var m = _marks[mi];
            if (!m || !m.path) continue;
            var el;
            try {
                var sel = '[data-path="' + m.path.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"]';
                el = _markRoot.querySelector(sel);
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

    // Register the coupling hook so Layer 1 can call us after DOM mutations
    window._watchReapplyHook = reapplyDecorations;

    // Public API exports
    window._officecliReapplyDecorations = reapplyDecorations;
    window._officecliApplyMarks = applyMarks;
    window._officecliSetMarks = function(arr) { _marks = arr || []; applyMarks(); };
    window._officecliGetMarks = function() { return _marks; };

    // ===== Click handler =====
    // Selects the closest element with [data-path].
    // Excel cells: shift = rectangular range from anchor, ctrl/cmd = toggle add.
    // Non-Excel elements: shift/ctrl/cmd = toggle multi-select.
    // Skipped if a rubber-band or cell drag just finished.
    var _suppressNextClick = false;
    var _lastInlineClickTime = 0;
    document.addEventListener('click', function(e) {
        if (_suppressNextClick || Date.now() - _lastInlineClickTime < 100) {
            _suppressNextClick = false; return;
        }
        var target = e.target.closest('[data-path]');
        if (!target) {
            // Don't clear selection when clicking UI chrome (sheet tabs, sidebar, etc.)
            if (e.target.closest('.sheet-tab, .sheet-tabs, .sidebar, .sidebar-toggle, .file-title, .page-counter, button, input, a')) return;
            if (!e.shiftKey && !e.ctrlKey && !e.metaKey && _selection.length > 0) {
                _selection = [];
                _anchor = null;
                applySelectionToDom();
                postSelection([]);
            }
            return;
        }
        var path = target.getAttribute('data-path');
        if (!path) return;
        var cell = _parseCellPath(path);

        if (e.shiftKey && _anchor && cell && cell.sheet === _anchor.sheet) {
            // Shift+click on Excel cell: select rectangular range from anchor
            _selection = _expandCellRange(_anchor.sheet, _anchor.col, _anchor.row, cell.col, cell.row);
        } else if ((e.ctrlKey || e.metaKey) && cell) {
            // Ctrl/Cmd+click on Excel cell: toggle individual cell
            var idx = _selection.indexOf(path);
            if (idx >= 0) _selection.splice(idx, 1);
            else { _selection.push(path); _anchor = cell; }
        } else if (e.shiftKey || e.ctrlKey || e.metaKey) {
            // Non-Excel element: toggle multi-select
            var idx = _selection.indexOf(path);
            if (idx >= 0) _selection.splice(idx, 1);
            else _selection.push(path);
        } else {
            // Plain click: select single, set anchor
            _selection = [path];
            if (cell) _anchor = cell;
        }
        applySelectionToDom(); // immediate visual feedback
        postSelection(_selection);
        e.preventDefault();
        e.stopPropagation();
    }, true);

    // ===== Chart drag-to-move =====
    var _chartDrag = null;
    // Expose drag-active flag so SSE full-update can defer body replacement
    window._isDragging = false;
    document.addEventListener('mousedown', function(e) {
        if (e.button !== 0) return;
        var chart = e.target.closest('.chart-container[data-path]');
        if (!chart) return;
        var path = chart.getAttribute('data-path');
        if (!path) return;
        _chartDrag = {
            el: chart, path: path,
            startX: e.clientX, startY: e.clientY,
            origLeft: chart.offsetLeft, origTop: chart.offsetTop,
            active: false
        };
        e.preventDefault();
    }, true);
    document.addEventListener('mousemove', function(e) {
        if (!_chartDrag) return;
        var dx = e.clientX - _chartDrag.startX;
        var dy = e.clientY - _chartDrag.startY;
        if (!_chartDrag.active) {
            if (Math.abs(dx) < 5 && Math.abs(dy) < 5) return;
            _chartDrag.active = true;
            window._isDragging = true;
            // Capture rect + width BEFORE placeholder insertion (flex reflow shifts position)
            var fixedRect = _chartDrag.el.getBoundingClientRect();
            var fixedW = _chartDrag.el.offsetWidth;
            _chartDrag.origFixedLeft = fixedRect.left;
            _chartDrag.origFixedTop = fixedRect.top;
            // Leave a dashed placeholder at original position
            var placeholder = document.createElement('div');
            placeholder.style.cssText = 'width:' + fixedW + 'px;height:' +
                _chartDrag.el.offsetHeight + 'px;border:2px dashed #217346;background:rgba(33,115,70,0.05);border-radius:4px;';
            _chartDrag.el.parentNode.insertBefore(placeholder, _chartDrag.el);
            _chartDrag.placeholder = placeholder;
            _chartDrag.el.style.position = 'fixed';
            _chartDrag.el.style.left = fixedRect.left + 'px';
            _chartDrag.el.style.top = fixedRect.top + 'px';
            _chartDrag.el.style.width = fixedW + 'px';
            _chartDrag.el.style.zIndex = '9999';
            _chartDrag.el.style.opacity = '0.7';
            _chartDrag.el.style.cursor = 'grabbing';
            _chartDrag.el.style.pointerEvents = 'none';
            _chartDrag.el.style.boxShadow = '0 4px 20px rgba(0,0,0,0.15)';
            _chartDrag.el.style.willChange = 'left, top';
        }
        _chartDrag.el.style.left = (_chartDrag.origFixedLeft + dx) + 'px';
        _chartDrag.el.style.top = (_chartDrag.origFixedTop + dy) + 'px';
    }, true);
    document.addEventListener('mouseup', function(e) {
        if (!_chartDrag) return;
        var cd = _chartDrag;
        _chartDrag = null;
        window._isDragging = false;
        if (!cd.active) return; // no drag, let click handle it
        // Reset visual + remove placeholder
        if (cd.placeholder) cd.placeholder.remove();
        cd.el.style.position = '';
        cd.el.style.zIndex = '';
        cd.el.style.opacity = '';
        cd.el.style.cursor = '';
        cd.el.style.pointerEvents = '';
        cd.el.style.left = '';
        cd.el.style.top = '';
        cd.el.style.width = '';
        cd.el.style.boxShadow = '';
        cd.el.style.willChange = '';
        var dx = e.clientX - cd.startX;
        var dy = e.clientY - cd.startY;
        if (Math.abs(dx) < 5 && Math.abs(dy) < 5) return;
        // Estimate row/col delta from pixel offset.
        // Average row height ≈ 20px, average col width ≈ 64px (from default Excel sizing).
        // Find actual average from visible row headers and col headers.
        var rowHeaders = document.querySelectorAll('.sheet-content.active th.row-header');
        var colHeaders = document.querySelectorAll('.sheet-content.active th.col-header');
        var avgRowH = 20, avgColW = 64;
        if (rowHeaders.length >= 2) {
            var first = rowHeaders[0].getBoundingClientRect();
            var last = rowHeaders[rowHeaders.length - 1].getBoundingClientRect();
            avgRowH = (last.top - first.top) / (rowHeaders.length - 1);
        }
        if (colHeaders.length >= 2) {
            var first = colHeaders[0].getBoundingClientRect();
            var last = colHeaders[colHeaders.length - 1].getBoundingClientRect();
            avgColW = (last.left - first.left) / (colHeaders.length - 1);
        }
        var dRows = Math.round(dy / Math.max(avgRowH, 1));
        var dCols = Math.round(dx / Math.max(avgColW, 1));
        if (dRows === 0 && dCols === 0) return;
        // Read current anchor from data-from-col/data-from-row on the overlay div
        var overlay = cd.el.closest('[data-from-col]') || cd.el.parentElement && cd.el.parentElement.closest('[data-from-col]');
        var currentCol = overlay ? parseInt(overlay.getAttribute('data-from-col'), 10) || 0 : 0;
        var currentRow = overlay ? parseInt(overlay.getAttribute('data-from-row'), 10) || 0 : 0;
        var newRow = Math.max(0, currentRow + dRows);
        var newCol = Math.max(0, currentCol + dCols);
        fetch('/api/edit', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: cd.path, props: {
                x: String(newCol),
                y: String(newRow)
            }})
        }).catch(function() {});
        _suppressNextClick = true;
    }, true);

    // ===== Double-click inline editing (Excel-style) =====
    var _editingCell = null; // currently editing td element
    document.addEventListener('dblclick', function(e) {
        var td = e.target.closest('td[data-path]');
        if (!td) return;
        var path = td.getAttribute('data-path');
        if (!path || !_parseCellPath(path)) return;
        if (_editingCell) return; // already editing

        _editingCell = td;
        var originalText = td.textContent || '';
        // Strip data-bar/icon overlays — get just the text node content
        // Show formula if cell has one, otherwise show displayed text
        var formula = td.getAttribute('data-formula');
        var textSpan = td.querySelector('.cell-text') || td;
        var editText = formula || textSpan.textContent || '';

        var input = document.createElement('input');
        input.type = 'text';
        input.value = editText;
        input.style.cssText = 'min-width:100%;height:100%;border:none;outline:2px solid #217346;' +
            'padding:1px 4px;font:inherit;background:#fff;box-sizing:border-box;' +
            'position:absolute;left:0;top:0;z-index:2000;white-space:nowrap;';
        td.style.position = 'relative';
        td.style.overflow = 'visible';
        td.appendChild(input);
        // Auto-expand width to fit content
        function autoSize() {
            input.style.width = '0';
            input.style.width = Math.max(td.offsetWidth, input.scrollWidth + 8) + 'px';
        }
        input.addEventListener('input', autoSize);
        input.focus();
        input.select();
        autoSize();

        function commit() {
            if (!_editingCell) return;
            var newValue = input.value;
            input.remove();
            _editingCell = null;
            if (newValue === editText) return; // no change
            // POST edit to watch server
            fetch('/api/edit', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ path: path, prop: 'text', value: newValue })
            }).catch(function() {});
        }
        function cancel() {
            if (!_editingCell) return;
            input.remove();
            _editingCell = null;
        }
        input.addEventListener('keydown', function(ke) {
            if (ke.key === 'Enter') { ke.preventDefault(); commit(); }
            else if (ke.key === 'Escape') { ke.preventDefault(); cancel(); }
        });
        input.addEventListener('blur', commit);
        e.preventDefault();
        e.stopPropagation();
    }, true);

    // ===== Cell-to-cell drag selection (Excel-style) =====
    // Mousedown on an Excel cell <td> starts a drag. Dragging to another cell
    // selects the rectangular range. Ctrl/Cmd+drag adds to existing selection.
    //
    // ===== Rubber-band (box) selection =====
    // Press on empty space (no [data-path] under cursor) and drag to draw a
    // selection rectangle. Any element whose bounding box intersects the
    // rectangle gets selected. Shift adds to current selection; plain replaces.
    // Esc cancels mid-drag.
    var _rubber = null; // {startX, startY, shift, div}
    var _RUBBER_THRESHOLD = 5; // px before treating as drag (vs click)

    document.addEventListener('mousedown', function(e) {
        if (e.button !== 0) return;

        // Excel header drag: drag on row/col headers to select multiple rows/columns
        var headerTh = e.target.closest('th.row-header[data-path], th.col-header[data-path]');
        if (headerTh) {
            var hpath = headerTh.getAttribute('data-path');
            var rowHdr = hpath.match(/^(\/[^/]+)\/row\[(\d+)\]$/);
            var colHdr = hpath.match(/^(\/[^/]+)\/col\[([A-Za-z]+)\]$/);
            if (rowHdr || colHdr) {
                _headerDrag = {
                    startX: e.clientX, startY: e.clientY,
                    type: rowHdr ? 'row' : 'col',
                    sheet: rowHdr ? rowHdr[1] : colHdr[1],
                    start: rowHdr ? parseInt(rowHdr[2], 10) : colHdr[2],
                    snapshot: _selection.slice(),
                    active: false
                };
                e.preventDefault();
                return;
            }
        }

        // Excel cell drag: start tracking on mousedown over a cell <td>
        var cellTd = e.target.closest('td[data-path]');
        if (cellTd) {
            var path = cellTd.getAttribute('data-path');
            var cell = _parseCellPath(path);
            if (cell) {
                var additive = e.ctrlKey || e.metaKey;
                _cellDrag = {
                    startX: e.clientX, startY: e.clientY,
                    anchor: cell,
                    base: additive ? _selection.slice() : [],
                    snapshot: _selection.slice(),
                    active: false
                };
                e.preventDefault();
                return;
            }
        }

        if (e.target.closest('[data-path]')) return; // non-cell data-path (PPT/Word)
        // Ignore mousedown inside scrollbars / sidebar / interactive UI
        if (e.target.closest('.sidebar, .sidebar-toggle, .page-counter, button, input, a')) return;
        _rubber = { startX: e.clientX, startY: e.clientY, shift: e.shiftKey, div: null };
    }, true);

    document.addEventListener('mousemove', function(e) {
        // Header drag (row/col)
        if (_headerDrag) {
            var dx = e.clientX - _headerDrag.startX;
            var dy = e.clientY - _headerDrag.startY;
            if (!_headerDrag.active) {
                if (Math.abs(dx) < _RUBBER_THRESHOLD && Math.abs(dy) < _RUBBER_THRESHOLD) return;
                _headerDrag.active = true;
                window._isDragging = true;
            }
            var el = document.elementFromPoint(e.clientX, e.clientY);
            if (el) {
                var th = el.closest('th[data-path]');
                if (th) {
                    var hp = th.getAttribute('data-path');
                    if (_headerDrag.type === 'row') {
                        var rm = hp.match(/^(\/[^/]+)\/row\[(\d+)\]$/);
                        if (rm && rm[1] === _headerDrag.sheet) {
                            var r1 = _headerDrag.start, r2 = parseInt(rm[2], 10);
                            var paths = [];
                            for (var r = Math.min(r1,r2); r <= Math.max(r1,r2); r++)
                                paths.push(_headerDrag.sheet + '/row[' + r + ']');
                            _selection = paths;
                            applySelectionToDom();
                        }
                    } else {
                        var cm = hp.match(/^(\/[^/]+)\/col\[([A-Za-z]+)\]$/);
                        if (cm && cm[1] === _headerDrag.sheet) {
                            var c1 = _colToNum(_headerDrag.start), c2 = _colToNum(cm[2]);
                            var paths = [];
                            for (var c = Math.min(c1,c2); c <= Math.max(c1,c2); c++)
                                paths.push(_headerDrag.sheet + '/col[' + _numToCol(c) + ']');
                            _selection = paths;
                            applySelectionToDom();
                        }
                    }
                }
            }
            return;
        }

        // Cell drag
        if (_cellDrag) {
            var dx = e.clientX - _cellDrag.startX;
            var dy = e.clientY - _cellDrag.startY;
            if (!_cellDrag.active) {
                if (Math.abs(dx) < _RUBBER_THRESHOLD && Math.abs(dy) < _RUBBER_THRESHOLD) return;
                _cellDrag.active = true;
                window._isDragging = true;
            }
            var el = document.elementFromPoint(e.clientX, e.clientY);
            if (el) {
                var td = el.closest('td[data-path]');
                if (td) {
                    var path = td.getAttribute('data-path');
                    var cell = _parseCellPath(path);
                    if (cell && cell.sheet === _cellDrag.anchor.sheet) {
                        var range = _expandCellRange(cell.sheet,
                            _cellDrag.anchor.col, _cellDrag.anchor.row,
                            cell.col, cell.row);
                        _selection = _uniquePaths(_cellDrag.base.concat(range));
                        applySelectionToDom(); // visual feedback only, no POST
                    }
                }
            }
            return;
        }

        // Rubber-band
        if (!_rubber) return;
        var dx = e.clientX - _rubber.startX;
        var dy = e.clientY - _rubber.startY;
        if (!_rubber.div) {
            if (Math.abs(dx) < _RUBBER_THRESHOLD && Math.abs(dy) < _RUBBER_THRESHOLD) return;
            var d = document.createElement('div');
            d.id = '_officecli_rubber';
            d.style.cssText = 'position:fixed;border:1.5px dashed #217346;' +
                'background:rgba(33,115,70,0.12);pointer-events:none;' +
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
        // Header drag commit
        if (_headerDrag) {
            var hd = _headerDrag;
            _headerDrag = null;
            window._isDragging = false;
            if (hd.active) {
                postSelection(_selection);
                _suppressNextClick = true;
                e.preventDefault();
                e.stopPropagation();
                return;
            }
            // Didn't drag — fall through to click handler
            return;
        }

        // Cell drag commit
        if (_cellDrag) {
            var cd = _cellDrag;
            _cellDrag = null;
            window._isDragging = false;
            if (cd.active) {
                // Drag completed — set anchor to drag start for future shift+clicks
                _anchor = cd.anchor;
                postSelection(_selection);
                _suppressNextClick = true;
                e.preventDefault();
                e.stopPropagation();
                return;
            }
            // Didn't move enough — handle click logic inline here because
            // mousedown's preventDefault() suppresses click for Ctrl (not Meta/Shift).
            var path = cd.anchor.sheet + '/' + cd.anchor.col + cd.anchor.row;
            var cell = cd.anchor;
            if (e.shiftKey && _anchor && cell.sheet === _anchor.sheet) {
                _selection = _expandCellRange(_anchor.sheet, _anchor.col, _anchor.row, cell.col, cell.row);
            } else if (e.ctrlKey || e.metaKey) {
                var idx = _selection.indexOf(path);
                if (idx >= 0) _selection.splice(idx, 1);
                else { _selection.push(path); _anchor = cell; }
            } else {
                _selection = [path];
                _anchor = cell;
            }
            applySelectionToDom(); // immediate visual feedback
            postSelection(_selection);
            // Suppress the click event that may follow (Meta/Shift are not
            // suppressed by mousedown's preventDefault on macOS).
            _lastInlineClickTime = Date.now();
            _suppressNextClick = true;
            return;
        }

        // Rubber-band commit
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

    function _cancelDrags() {
        if (_rubber) { if (_rubber.div) _rubber.div.remove(); _rubber = null; }
        if (_cellDrag) {
            _selection = _cellDrag.snapshot.slice();
            applySelectionToDom();
            _cellDrag = null;
            _suppressNextClick = true;
        }
        if (_headerDrag) {
            _selection = _headerDrag.snapshot.slice();
            applySelectionToDom();
            _headerDrag = null;
            _suppressNextClick = true;
        }
        window._isDragging = false;
    }
    var _cancelRubber = _cancelDrags;

    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') _cancelDrags();
    });

    // If the user alt-tabs / window loses focus mid-drag, the OS-level
    // mouseup never reaches us. Clean up so the rubber-band overlay
    // doesn't get stuck on screen and click handling stays sane.
    window.addEventListener('blur', _cancelDrags);
    document.addEventListener('visibilitychange', function() {
        if (document.hidden) _cancelDrags();
    });
    // Belt-and-suspenders: if a mouseup never came after a long enough
    // mousemove pause, drop the rubber-band on the next mouse re-entry.
    document.addEventListener('mouseleave', function(e) {
        // Only cancel if cursor truly left the page (relatedTarget == null)
        if (!e.relatedTarget && _rubber) _cancelDrags();
    });

    // ===== SSE: selection and mark metadata updates =====
    if (es) {
        es.addEventListener('update', function(e) {
            var msg;
            try { msg = JSON.parse(e.data); } catch (err) { return; }
            if (msg.action === 'selection-update') {
                var newPaths = msg.paths || [];
                // Skip re-apply if selection unchanged (avoids flicker when
                // SSE echoes back the same selection we just set locally)
                if (JSON.stringify(newPaths) === JSON.stringify(_selection)) return;
                _selection = newPaths;
                applySelectionToDom();
            } else if (msg.action === 'mark-update') {
                // Monotonic version: clients may CAS on this value to skip
                // redundant updates if they missed nothing. We just refresh.
                _marks = msg.marks || [];
                applyMarks();
            }
        });
    }
})();
