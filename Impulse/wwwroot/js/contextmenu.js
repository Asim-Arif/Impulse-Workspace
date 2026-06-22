window.blazorContextMenuHelper = {
    init: function (dotNetHelper) {
        // Keep track of scroll position so menu stays in correct spot
        window.addEventListener('scroll', () => {
            dotNetHelper.invokeMethodAsync('UpdateScrollPosition', window.scrollX, window.scrollY);
        });
    }
};

window.initializeContextMenus = () => {
    if (BlazorContextMenu && BlazorContextMenu.initialize) {
        BlazorContextMenu.initialize();

        // Register menus dynamically
        BlazorContextMenu.register('#item1', [
            { text: 'Option 1', callback: () => alert('Item 1 - Option 1') },
            { text: 'Option 2', callback: () => alert('Item 1 - Option 2') }
        ]);

        BlazorContextMenu.register('#item2', [
            { text: 'Option A', callback: () => alert('Item 2 - Option A') },
            { text: 'Option B', callback: () => alert('Item 2 - Option B') }
        ]);
    }
};