window.getWindowWidth = function () {
    return window.innerWidth;
};

// Automatic native tooltips on hover for all sidebar navigation links and headers
document.addEventListener('mouseover', function (e) {
    var link = e.target.closest('.sidebar .nav-link, .sidebar .nav-section-header, .sidebar .nav-sub-header');
    if (link && !link.getAttribute('title')) {
        var text = (link.textContent || link.innerText || '').trim();
        if (text) {
            link.setAttribute('title', text);
        }
    }
});

// ── Global Enter-as-Tab Navigation for Rapid Data Entry ─────────────────────
document.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter') return;

    // Check if the focused element is within an enter-as-tab container
    var container = e.target.closest('.enter-as-tab, [data-enter-as-tab="true"]');
    if (!container) return;

    var target = e.target;

    // Allow normal Enter in multi-line textareas (unless Ctrl is pressed)
    if (target.tagName === 'TEXTAREA' && !e.ctrlKey) {
        return;
    }

    // Allow Enter on buttons to trigger their action naturally
    if (target.tagName === 'BUTTON' || (target.tagName === 'INPUT' && (target.type === 'button' || target.type === 'submit'))) {
        return;
    }

    // If target is inside a BlazoredTypeahead with open dropdown results, let the typeahead select the item
    var isTypeahead = target.classList.contains('blazored-typeahead__input');
    if (isTypeahead) {
        var dropdownMenu = target.closest('.blazored-typeahead')?.querySelector('.blazored-typeahead__results');
        if (dropdownMenu && dropdownMenu.offsetParent !== null && dropdownMenu.children.length > 0) {
            return;
        }
    }

    e.preventDefault();

    // Find all visible, enabled, non-readonly focusable elements within the form/container
    var selector = 'input:not([type="hidden"]):not([disabled]):not([readonly]):not([tabindex="-1"]), ' +
                   'select:not([disabled]):not([readonly]):not([tabindex="-1"]), ' +
                   'textarea:not([disabled]):not([readonly]):not([tabindex="-1"]), ' +
                   'button:not([disabled]):not([tabindex="-1"])';

    var focusables = Array.prototype.filter.call(container.querySelectorAll(selector), function (el) {
        return el.offsetParent !== null && window.getComputedStyle(el).visibility !== 'hidden';
    });

    var index = focusables.indexOf(target);
    if (index === -1) return;

    var nextIndex = e.shiftKey ? index - 1 : index + 1;
    if (nextIndex >= 0 && nextIndex < focusables.length) {
        var nextEl = focusables[nextIndex];
        nextEl.focus();
        if (typeof nextEl.select === 'function' && nextEl.tagName === 'INPUT' && nextEl.type !== 'checkbox' && nextEl.type !== 'radio' && nextEl.type !== 'date') {
            nextEl.select();
        }
    }
});