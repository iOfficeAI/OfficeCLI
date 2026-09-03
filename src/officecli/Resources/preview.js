// OfficeCli HTML Preview Script
(function() {
    const main = document.querySelector('.main');
    const sidebar = document.querySelector('.sidebar');
    const counter = document.querySelector('.page-counter');
    let currentSlide = 0;
    let isFullscreen = false;

    // ===== Live DOM queries (SSE may add/remove elements) =====
    function getContainers() { return [...document.querySelectorAll('.main > .slide-container')]; }
    function getThumbs() { return [...document.querySelectorAll('.sidebar > .thumb')]; }
    function getTotal() { return getContainers().length; }

    // ===== Responsive scaling =====
    function scaleSlides() {
        // 40px breathing room for interactive viewing; headless captures fill the
        // viewport (the screenshot path sizes it to the slide) so they take none.
        const headless = document.documentElement.classList.contains('headless');
        // For an exact headless capture, use the browser's emulated viewport
        // directly. Flex children can report a narrower client box while the
        // screenshot canvas is still the full viewport (especially on Windows
        // Edge/Chrome), which otherwise shrinks a 16:9 slide to ~1116x627.
        const availW = headless ? window.innerWidth : main.clientWidth - 40;
        const availH = headless ? window.innerHeight : main.clientHeight - 40;
        const slides = document.querySelectorAll('.main > .slide-container .slide');
        // A lone headless slide is a single-slide screenshot: scale it to fill the
        // viewport in BOTH directions (up or down) so the capture is flush at the
        // requested resolution regardless of the slide's physical size. Interactive
        // and multi-slide views only ever shrink to fit.
        const fill = headless && slides.length === 1;
        // Chromium's command-line screenshot can expose a layout viewport that
        // is shorter than the bitmap it will emit (window chrome is included in
        // --window-size).  For the one-slide capture path, use the outer canvas
        // height when it is larger so the slide fills the requested bitmap
        // instead of shrinking into a dark letterbox.  This is deliberately
        // limited to headless one-slide previews: scrolling/grid/range previews
        // and non-PPT HTML retain their normal responsive layout.
        if (fill) {
            const canvasH = Math.max(window.innerHeight, window.outerHeight || 0);
            if (canvasH > main.clientHeight) {
                document.documentElement.style.height = canvasH + 'px';
                document.body.style.height = canvasH + 'px';
                main.style.height = canvasH + 'px';
                main.style.flex = 'none';
            }
        }
        const captureH = fill ? main.clientHeight : availH;
        slides.forEach(slide => {
            const designW = slide.offsetWidth;
            if (availW > 0 && captureH > 0 && (fill || designW > availW)) {
                // A headless single-slide capture must fit both layout axes.
                // Chromium's legacy command-line screenshot can expose a
                // shorter layout viewport than its requested bitmap; fitting by
                // height preserves footers/legends instead of cutting them off.
                const s = fill
                    ? Math.min(availW / designW, captureH / slide.offsetHeight)
                    : availW / designW;
                slide.style.transform = `scale(${s})`;
                slide.style.transformOrigin = 'center top';
                const designH = slide.offsetHeight;
                slide.parentElement.style.height = (designH * s) + 'px';
                slide.parentElement.style.width = (designW * s) + 'px';
            } else {
                slide.style.transform = '';
                slide.parentElement.style.height = '';
                slide.parentElement.style.width = '';
            }
        });
    }
    scaleSlides();
    window.scaleSlides = scaleSlides;
    window.addEventListener('resize', scaleSlides);

    // ===== Sidebar thumbnails =====
    function setActiveThumb(idx) {
        getThumbs().forEach((t, i) => t.classList.toggle('active', i === idx));
        currentSlide = idx;
        if (counter) counter.textContent = `${idx + 1} / ${getTotal()}`;
    }

    // Event delegation for thumb clicks (handles SSE-added thumbs)
    if (sidebar) {
        sidebar.addEventListener('click', function(e) {
            const thumb = e.target.closest('.thumb');
            if (!thumb) return;
            const thumbs = getThumbs();
            const idx = thumbs.indexOf(thumb);
            if (idx < 0) return;
            if (isFullscreen) { showFullscreenSlide(idx); return; }
            const containers = getContainers();
            if (containers[idx]) {
                containers[idx].scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
            setActiveThumb(idx);
        });
    }

    // Track visible slide on scroll (normal mode) — use MutationObserver to auto-observe new slides
    let scrollObserver;
    if (main) {
        scrollObserver = new IntersectionObserver(entries => {
            if (isFullscreen) return;
            const containers = getContainers();
            entries.forEach(e => {
                if (e.isIntersecting && e.intersectionRatio > 0.3) {
                    const idx = containers.indexOf(e.target);
                    if (idx >= 0) setActiveThumb(idx);
                }
            });
        }, { root: main, threshold: 0.3 });
        getContainers().forEach(c => scrollObserver.observe(c));

        // Auto-observe new slide-containers added to main
        new MutationObserver(mutations => {
            mutations.forEach(m => {
                m.addedNodes.forEach(node => {
                    if (node.nodeType === 1 && node.classList.contains('slide-container')) {
                        scrollObserver.observe(node);
                    }
                });
            });
        }).observe(main, { childList: true });
    }

    // ===== Fullscreen mode =====
    function showFullscreenSlide(idx) {
        const containers = getContainers();
        const total = containers.length;
        idx = Math.max(0, Math.min(idx, total - 1));
        containers.forEach((c, i) => c.classList.toggle('fs-active', i === idx));
        setActiveThumb(idx);
        const slide = containers[idx]?.querySelector('.slide');
        if (slide) {
            const vw = window.innerWidth, vh = window.innerHeight - 30;
            const sw = slide.scrollWidth || slide.offsetWidth;
            const sh = slide.scrollHeight || slide.offsetHeight;
            const s = Math.min(vw / sw, vh / sh, 1);
            slide.style.transform = `scale(${s})`;
            slide.style.transformOrigin = 'center top';
        }
    }
    function enterFullscreen() {
        isFullscreen = true;
        document.body.classList.add('fullscreen');
        showFullscreenSlide(currentSlide);
    }
    function exitFullscreen() {
        isFullscreen = false;
        document.body.classList.remove('fullscreen');
        getContainers().forEach(c => { c.classList.remove('fs-active'); c.style.display = ''; });
        scaleSlides();
        getContainers()[currentSlide]?.scrollIntoView({ block: 'center' });
    }

    // ===== Keyboard navigation =====
    document.addEventListener('keydown', e => {
        if (e.key === 'f' || e.key === 'F') {
            e.preventDefault();
            isFullscreen ? exitFullscreen() : enterFullscreen();
            return;
        }
        if (e.key === 'Escape' && isFullscreen) {
            e.preventDefault();
            exitFullscreen();
            return;
        }
        const next = e.key === 'ArrowDown' || e.key === ' ' || e.key === 'ArrowRight';
        const prev = e.key === 'ArrowUp' || e.key === 'ArrowLeft';
        if (!next && !prev) return;
        e.preventDefault();

        const total = getTotal();
        if (isFullscreen) {
            showFullscreenSlide(currentSlide + (next ? 1 : -1));
        } else {
            const target = next
                ? Math.min(currentSlide + 1, total - 1)
                : Math.max(currentSlide - 1, 0);
            const containers = getContainers();
            if (containers[target]) {
                containers[target].scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
            setActiveThumb(target);
        }
    });

    // ===== Populate & scale thumbnail slides via cloneNode (zero base64 duplication) =====
    function buildThumbs() {
        const slides = document.querySelectorAll('.main > .slide-container .slide');
        const inners = document.querySelectorAll('.thumb-inner');
        slides.forEach((slide, i) => {
            if (i >= inners.length) return;
            const inner = inners[i];
            if (inner.querySelector('.thumb-slide')) return;
            const clone = slide.cloneNode(true);
            clone.className = 'thumb-slide';
            clone.style.transform = '';
            // Remove IDs from cloned elements to avoid getElementById conflicts (e.g. 3D canvas)
            clone.querySelectorAll('[id]').forEach(el => el.removeAttribute('id'));
            // Remove cloned <script> tags (module scripts won't re-execute but keep DOM clean)
            clone.querySelectorAll('script').forEach(el => el.remove());
            inner.appendChild(clone);
        });
        scaleThumbs();
    }
    function scaleThumbs() {
        document.querySelectorAll('.thumb-inner').forEach(inner => {
            const thumbSlide = inner.querySelector('.thumb-slide');
            if (!thumbSlide) return;
            const thumbW = inner.clientWidth;
            const slideW = thumbSlide.scrollWidth || thumbSlide.offsetWidth;
            if (slideW > 0 && thumbW > 0) {
                thumbSlide.style.transform = `scale(${thumbW / slideW})`;
                thumbSlide.style.transformOrigin = '0 0';
            }
        });
    }
    buildThumbs();
    window.buildThumbs = buildThumbs;
    window.scaleThumbs = scaleThumbs;
    window.addEventListener('resize', scaleThumbs);

    // ===== Sidebar toggle (exposed globally for onclick) =====
    window.toggleSidebar = function() {
        document.body.classList.toggle('sidebar-visible');
        document.body.classList.toggle('sidebar-hidden');
        // Re-scale after sidebar toggle changes main area width
        requestAnimationFrame(() => {
            scaleSlides();
            buildThumbs();
            scaleThumbs();
        });
    };

    // Init
    if (getTotal() > 0) setActiveThumb(0);
})();
