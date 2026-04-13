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

    var _SEL_CLASSES = ['officecli-selected', 'officecli-sel-range',
        'officecli-sel-t', 'officecli-sel-b', 'officecli-sel-l', 'officecli-sel-r',
        'officecli-sel-handle'];

    function applySelectionToDom() {
        // Clear all selection classes
        var allSel = _SEL_CLASSES.map(function(c) { return '.' + c; }).join(',');
        document.querySelectorAll(allSel).forEach(function(el) {
            _SEL_CLASSES.forEach(function(c) { el.classList.remove(c); });
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
            // Apply range fill + edge borders to cells
            for (var i = 0; i < rect.cells.length; i++) {
                var cell = rect.cells[i];
                try {
                    var sel = '[data-path="' + cell.path.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"]';
                    document.querySelectorAll(sel).forEach(function(el) {
                        el.classList.add('officecli-sel-range');
                        if (cell.row === rect.minR) el.classList.add('officecli-sel-t');
                        if (cell.row === rect.maxR) el.classList.add('officecli-sel-b');
                        if (cell.col === rect.minC) el.classList.add('officecli-sel-l');
                        if (cell.col === rect.maxC) el.classList.add('officecli-sel-r');
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
                    // Row header: highlight row cells
                    var rowMatch = path.match(/^(\/[^/]+)\/row\[(\d+)\]$/);
                    if (rowMatch && el.tagName === 'TH') {
                        var tr = el.closest('tr');
                        if (tr) tr.querySelectorAll('td[data-path]').forEach(function(td) {
                            td.classList.add('officecli-sel-range');
                        });
                    }
                    // Col header: highlight column cells
                    var colMatch = path.match(/^(\/[^/]+)\/col\[([A-Za-z]+)\]$/);
                    if (colMatch && el.tagName === 'TH') {
                        var sheet = colMatch[1], col = colMatch[2];
                        var re = new RegExp('^' + sheet.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\/' + col + '\\d+$', 'i');
                        document.querySelectorAll('td[data-path]').forEach(function(td) {
                            if (re.test(td.getAttribute('data-path')))
                                td.classList.add('officecli-sel-range');
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
            // Range fill: light gray/blue like real Excel
            'td.officecli-sel-range{' +
                'background:rgba(33,115,70,0.10) !important;' +
                'position:relative;' +
            '}' +
            // Range perimeter borders (only on edges)
            'td.officecli-sel-t{border-top:2px solid #217346 !important;}' +
            'td.officecli-sel-b{border-bottom:2px solid #217346 !important;}' +
            'td.officecli-sel-l{border-left:2px solid #217346 !important;}' +
            'td.officecli-sel-r{border-right:2px solid #217346 !important;}' +
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
    document.addEventListener('click', function(e) {
        if (_suppressNextClick) { _suppressNextClick = false; return; }
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
            if (cd.active) {
                // Drag completed — set anchor to drag start for future shift+clicks
                _anchor = cd.anchor;
                postSelection(_selection);
                _suppressNextClick = true;
                e.preventDefault();
                e.stopPropagation();
                return;
            }
            // Didn't move enough — mousedown's preventDefault() suppressed the
            // click event, so handle click logic (Ctrl/Shift/plain) inline here.
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
            _suppressNextClick = true;
            applySelectionToDom(); // immediate visual feedback
            postSelection(_selection);
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
    }

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
                _selection = msg.paths || [];
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
