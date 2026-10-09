const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(
    path.join(__dirname, '../src/officecli/Resources/watch-sse-core.js'), 'utf8');

function viewer() {
    const navigations = [];
    const timers = new Map();
    let timerId = 0;
    const slides = [];
    const main = {
        appendChild(slide) {
            slide.parentNode = main;
            slides.push(slide);
        },
        replaceChild(replacement, original) {
            replacement.parentNode = main;
            slides[slides.indexOf(original)] = replacement;
        },
    };
    function slide(number, html) {
        return {
            number, html,
            querySelectorAll() { return []; },
            scrollIntoView() { navigations.push(number); },
        };
    }
    for (let number = 1; number <= 3; number++) {
        main.appendChild(slide(number, 'original'));
    }
    const document = {
        querySelector(selector) {
            if (selector === '.main') return main;
            const match = selector.match(/data-slide="(\d+)"/);
            return match ? slides.find(slide => slide.number === Number(match[1])) : null;
        },
        createElement() {
            return {
                set innerHTML(html) {
                    const number = Number(html.match(/data-slide="(\d+)"/)[1]);
                    this.firstElementChild = slide(number, html);
                },
            };
        },
    };
    class EventSource extends EventTarget {}
    const window = {};
    vm.runInNewContext(source, {
        EventSource, window, document,
        setTimeout(callback) { timers.set(++timerId, callback); return timerId; },
        clearTimeout(id) { timers.delete(id); },
        location: { reload() { assert.fail('unexpected reload'); } },
    });
    return {
        navigations, slides,
        update(message) {
            window._watchEs.dispatchEvent(new MessageEvent('update', {
                data: JSON.stringify(message),
            }));
            for (const callback of timers.values()) callback();
            timers.clear();
        },
    };
}

test('replacing a slide updates its content without navigating the viewer', () => {
    const page = viewer();
    const html = '<div class="slide-container" data-slide="3">edited</div>';
    page.update({ action: 'replace', slide: 3, html });
    assert.equal(page.slides[2].html, html);
    assert.deepEqual(page.navigations, []);
});

test('adding a slide updates the document without navigating the viewer', () => {
    const page = viewer();
    const html = '<div class="slide-container" data-slide="4">new slide</div>';
    page.update({ action: 'add', slide: 4, html });
    assert.equal(page.slides.length, 4);
    assert.equal(page.slides[3].html, html);
    assert.deepEqual(page.navigations, []);
});

test('explicit slide navigation still scrolls to its target', () => {
    const page = viewer();
    page.update({ action: 'scroll', scrollTo: '.slide-container[data-slide="3"]' });
    assert.deepEqual(page.navigations, [3]);
});
