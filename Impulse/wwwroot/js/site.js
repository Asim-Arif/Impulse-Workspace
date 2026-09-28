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