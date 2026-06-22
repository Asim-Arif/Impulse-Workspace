function resetFileInput(element) {
    element.value = '';
}
window.focusElement = (id) => {
    const el = document.getElementById(id);
    if (el) {
        setTimeout(() => {
            el.focus();
            if (typeof el.select === "function") {
                el.select();
            }
        }, 10);
    }
};