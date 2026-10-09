/**
 * Impulse ERP - Cross-Tab Authentication Synchronization
 * Ensures all browser tabs stay in 100% sync when a user logs in, logs out, or switches accounts.
 */
(function () {
    const CHANNEL_NAME = 'impulse_auth_channel';
    const STORAGE_EVENT_KEY = 'impulse_auth_sync_event';
    const STORAGE_ACTIVE_USER_KEY = 'impulse_active_username';

    let broadcastChannel = null;

    try {
        if (typeof BroadcastChannel !== 'undefined') {
            broadcastChannel = new BroadcastChannel(CHANNEL_NAME);
        }
    } catch (e) {
        console.warn('[AuthSync] BroadcastChannel not supported:', e);
    }

    /**
     * Broadcast an authentication state change across all open tabs.
     * @param {'login' | 'logout' | 'user_changed'} action 
     * @param {string} userName 
     */
    window.notifyAuthChange = function (action, userName) {
        const payload = {
            action: action,
            userName: userName || '',
            timestamp: Date.now()
        };

        if (action === 'login' || action === 'user_changed') {
            try {
                localStorage.setItem(STORAGE_ACTIVE_USER_KEY, userName || '');
            } catch (e) { }
        } else if (action === 'logout') {
            try {
                localStorage.removeItem(STORAGE_ACTIVE_USER_KEY);
            } catch (e) { }
        }

        // 1. Send via BroadcastChannel (0ms instant cross-tab notification)
        if (broadcastChannel) {
            try {
                broadcastChannel.postMessage(payload);
            } catch (e) { }
        }

        // 2. Send via localStorage storage event (fallback across tabs)
        try {
            localStorage.setItem(STORAGE_EVENT_KEY, JSON.stringify(payload));
        } catch (e) { }
    };

    /**
     * Initialize this tab's user identity tracking
     * @param {string} tabUserName 
     */
    window.initTabAuthSync = function (tabUserName) {
        window.__impulse_tab_user = tabUserName || '';

        if (tabUserName) {
            try {
                localStorage.setItem(STORAGE_ACTIVE_USER_KEY, tabUserName);
            } catch (e) { }
        }
    };

    /**
     * Handle auth change events received from another tab
     */
    function handleAuthEvent(eventData) {
        if (!eventData || !eventData.action) return;

        const currentPath = (window.location.pathname || '').toLowerCase();
        const isAuthPage = currentPath.includes('/identity/account/login') || currentPath.includes('/identity/account/logout');

        if (eventData.action === 'logout') {
            if (!isAuthPage) {
                console.warn('[AuthSync] Logout detected in another tab. Redirecting to login.');
                const returnUrl = encodeURIComponent(window.location.pathname + window.location.search);
                window.location.replace('/Identity/Account/Login?returnUrl=' + returnUrl);
            }
        } else if (eventData.action === 'login' || eventData.action === 'user_changed') {
            if (!isAuthPage) {
                const currentTabUser = (window.__impulse_tab_user || '').toLowerCase();
                const newGlobalUser = (eventData.userName || '').toLowerCase();

                // If the user changed to a different account, reload this tab immediately
                if (currentTabUser !== newGlobalUser) {
                    console.warn('[AuthSync] User change detected in another tab (' + currentTabUser + ' -> ' + newGlobalUser + '). Reloading tab.');
                    window.location.reload();
                }
            } else if (eventData.action === 'login') {
                // If on login page and another tab logged in, redirect to root
                window.location.replace('/');
            }
        }
    }

    // 1. Listen on BroadcastChannel
    if (broadcastChannel) {
        broadcastChannel.onmessage = function (event) {
            handleAuthEvent(event.data);
        };
    }

    // 2. Listen on storage events (cross-tab fallback)
    window.addEventListener('storage', function (event) {
        if (event.key === STORAGE_EVENT_KEY && event.newValue) {
            try {
                const data = JSON.parse(event.newValue);
                handleAuthEvent(data);
            } catch (e) { }
        }
    });

    // 3. Tab Focus / Visibility Re-check (handles background or sleeping tabs)
    function checkCurrentSessionOnFocus() {
        const currentPath = (window.location.pathname || '').toLowerCase();
        const isAuthPage = currentPath.includes('/identity/account/login') || currentPath.includes('/identity/account/logout');
        if (isAuthPage) return;

        try {
            const activeUser = localStorage.getItem(STORAGE_ACTIVE_USER_KEY);
            const currentTabUser = window.__impulse_tab_user;

            // If active user is cleared (logged out) or different user logged in
            if (activeUser === null && currentTabUser) {
                console.warn('[AuthSync] Session ended in another tab while in background. Redirecting to login.');
                window.location.replace('/Identity/Account/Login');
            } else if (activeUser && currentTabUser && activeUser.toLowerCase() !== currentTabUser.toLowerCase()) {
                console.warn('[AuthSync] Active user changed to ' + activeUser + ' while in background. Reloading tab.');
                window.location.reload();
            }
        } catch (e) { }
    }

    window.addEventListener('focus', checkCurrentSessionOnFocus);
    document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible') {
            checkCurrentSessionOnFocus();
        }
    });
})();
