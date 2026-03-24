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
        const availW = main.clientWidth - 40;
        document.querySelectorAll('.main > .slide-container .slide').forEach(slide => {
            const designW = slide.offsetWidth;
            if (designW > availW && availW > 0) {
                const s = availW / designW;
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
