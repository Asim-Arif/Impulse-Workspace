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
(function () {
    function getFocusableElements(container) {
        var selector = 'input:not([type="hidden"]):not([disabled]):not([readonly]):not([tabindex="-1"]), ' +
                       'select:not([disabled]):not([readonly]):not([tabindex="-1"]), ' +
                       'textarea:not([disabled]):not([readonly]):not([tabindex="-1"]), ' +
                       'button:not([disabled]):not([tabindex="-1"]):not(.blazored-typeahead__toggle-dropdown):not(.blazored-typeahead__clear-icon), ' +
                       '.blazored-typeahead__input-mask:not([tabindex="-1"])';

        var raw = container.querySelectorAll(selector);
        var focusables = [];

        for (var i = 0; i < raw.length; i++) {
            var el = raw[i];
            if (el.offsetParent === null || window.getComputedStyle(el).visibility === 'hidden') {
                continue;
            }
            if (el.closest('.blazored-typeahead--disabled, .disabled-typeahead-wrapper')) {
                continue;
            }
            // If inside a typeahead and its mask is already visible, do not include the underlying input
            if (el.classList.contains('blazored-typeahead__input')) {
                var mask = el.closest('.blazored-typeahead')?.querySelector('.blazored-typeahead__input-mask');
                if (mask && mask.offsetParent !== null) {
                    continue;
                }
            }
            focusables.push(el);
        }
        return focusables;
    }

    function advanceFocus(container, currentControl, isShift) {
        var focusables = getFocusableElements(container);
        var index = focusables.indexOf(currentControl);

        if (index === -1) {
            var typeahead = currentControl.closest('.blazored-typeahead');
            if (typeahead) {
                var activeInside = typeahead.querySelector('.blazored-typeahead__input-mask, .blazored-typeahead__input');
                index = focusables.indexOf(activeInside);
            }
        }
        if (index === -1) return;

        var nextIndex = isShift ? index - 1 : index + 1;
        if (nextIndex >= 0 && nextIndex < focusables.length) {
            var nextEl = focusables[nextIndex];
            nextEl.focus();
            if (typeof nextEl.select === 'function' && nextEl.tagName === 'INPUT' && nextEl.type !== 'checkbox' && nextEl.type !== 'radio' && nextEl.type !== 'date') {
                nextEl.select();
            }
        }
    }

    // Intercept BlazoredTypeahead.setFocus to support deterministic 1-Step auto-advance in Blazor Server
    if (window.blazoredTypeahead) {
        var origSetFocus = window.blazoredTypeahead.setFocus;
        window.blazoredTypeahead.setFocus = function (element) {
            var typeaheadEl = element?.closest('.blazored-typeahead');
            if (typeaheadEl && typeaheadEl._advanceOnSelect) {
                typeaheadEl._advanceOnSelect = false;
                var container = typeaheadEl.closest('.enter-as-tab, [data-enter-as-tab="true"]');
                if (container) {
                    advanceFocus(container, element, false);
                    return;
                }
            }
            if (origSetFocus) {
                origSetFocus.apply(this, arguments);
            } else if (element && element.focus) {
                element.focus();
            }
        };
    }

    document.addEventListener('keydown', function (e) {
        if (e.key !== 'Enter') return;

        var container = e.target.closest('.enter-as-tab, [data-enter-as-tab="true"]');
        if (!container) return;

        var target = e.target;

        // Allow normal Enter in multi-line textareas (unless Ctrl is pressed)
        if (target.tagName === 'TEXTAREA' && !e.ctrlKey) {
            return;
        }

        // Allow action buttons to trigger their action naturally (excluding typeahead helper buttons)
        if ((target.tagName === 'BUTTON' && !target.classList.contains('blazored-typeahead__toggle-dropdown') && !target.classList.contains('blazored-typeahead__clear-icon')) ||
            (target.tagName === 'INPUT' && (target.type === 'button' || target.type === 'submit'))) {
            return;
        }

        var currentControl = target.closest(
            '.blazored-typeahead__input, .blazored-typeahead__input-mask, input, select, textarea, button'
        ) || target;

        var typeaheadEl = target.closest('.blazored-typeahead');
        var dropdownMenu = typeaheadEl?.querySelector('.blazored-typeahead__results');
        var isDropdownOpen = dropdownMenu && dropdownMenu.offsetParent !== null && dropdownMenu.children.length > 0;

        if (typeaheadEl && isDropdownOpen) {
            var is1Step = container.getAttribute('data-enter-1step') === 'true';
            if (is1Step) {
                typeaheadEl._advanceOnSelect = true;
                // Safety timeout in case Blazor setFocus is bypassed
                setTimeout(function () {
                    if (typeaheadEl._advanceOnSelect) {
                        typeaheadEl._advanceOnSelect = false;
                        advanceFocus(container, currentControl, false);
                    }
                }, 150);
            }
            // Let the Enter event bubble so BlazoredTypeahead handles item selection
            return;
        }

        e.preventDefault();
        advanceFocus(container, currentControl, e.shiftKey);
    });
})();

// ── Table Keyboard Navigation Helper (Arrow Up/Down, Space, Home, End) ───────
document.addEventListener('keydown', function (e) {
    var tableNav = e.target.closest('[data-keyboard-nav="true"]');
    if (!tableNav) return;

    // Do not prevent default if focused on an editable form input
    if (['INPUT', 'SELECT', 'TEXTAREA'].indexOf(e.target.tagName) !== -1) {
        return;
    }

    if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === ' ' || e.key === 'Home' || e.key === 'End') {
        e.preventDefault();
    }
});

window.scrollElementIntoView = function (elementId) {
    var el = document.getElementById(elementId);
    if (el) {
        el.scrollIntoView({ block: 'nearest' });
    }
};