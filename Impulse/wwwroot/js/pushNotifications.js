// Impulse ERP Web Push Client Integration
window.impulsePush = {
    _vapidPublicKey: null,

    isSupported: function () {
        return ('serviceWorker' in navigator) && ('Notification' in window);
    },

    getPermissionState: function () {
        if (!('Notification' in window)) return 'unsupported';
        return Notification.permission;
    },

    urlB64ToUint8Array: function (base64String) {
        const padding = '='.repeat((4 - base64String.length % 4) % 4);
        const base64 = (base64String + padding)
            .replace(/\-/g, '+')
            .replace(/_/g, '/');

        const rawData = window.atob(base64);
        const outputArray = new Uint8Array(rawData.length);

        for (let i = 0; i < rawData.length; ++i) {
            outputArray[i] = rawData.charCodeAt(i);
        }
        return outputArray;
    },

    getApiBaseUrl: function () {
        const baseEl = document.querySelector('base');
        let baseHref = baseEl ? (baseEl.href || baseEl.getAttribute('href') || '/') : '/';
        if (!baseHref.endsWith('/')) baseHref += '/';
        return baseHref;
    },

    init: async function (userId, userName) {
        if (!this.isSupported()) {
            console.warn('[ImpulsePush] Push notifications are not supported in this browser.');
            return false;
        }

        try {
            const apiBase = this.getApiBaseUrl();
            const swUrl = apiBase + 'service-worker.js';

            // Register service worker
            const swRegistration = await navigator.serviceWorker.register(swUrl, { scope: apiBase });
            await navigator.serviceWorker.ready;

            // If permission is already granted, ensure subscription is synced with current user
            if (Notification.permission === 'granted') {
                return await this._syncSubscription(swRegistration, userId, userName);
            }

            return false;
        } catch (error) {
            console.error('[ImpulsePush] Error initializing push service worker:', error);
            return false;
        }
    },

    requestPermissionAndSubscribe: async function (userId, userName) {
        if (!('Notification' in window)) {
            return { success: false, reason: 'Notification API unavailable (requires secure context)' };
        }
        if (!('serviceWorker' in navigator)) {
            return { success: false, reason: 'ServiceWorker API unavailable' };
        }

        try {
            const permission = await Notification.requestPermission();
            if (permission !== 'granted') {
                return { success: false, reason: 'Permission is ' + permission };
            }

            const apiBase = this.getApiBaseUrl();
            const swUrl = apiBase + 'service-worker.js';
            const swRegistration = await navigator.serviceWorker.register(swUrl, { scope: apiBase });
            await navigator.serviceWorker.ready;

            const subscribed = await this._syncSubscription(swRegistration, userId, userName);
            return { success: subscribed, reason: subscribed ? 'Subscribed successfully' : 'Failed to register with server' };
        } catch (error) {
            console.error('[ImpulsePush] Error requesting push permission:', error);
            return { success: false, reason: (error.name || 'Error') + ': ' + (error.message || error) };
        }
    },

    _syncSubscription: async function (swRegistration, userId, userName) {
        try {
            const apiBase = this.getApiBaseUrl();

            // 1. Fetch public key if not cached
            if (!this._vapidPublicKey) {
                const response = await fetch(apiBase + 'api/push/vapid-public-key');
                if (!response.ok) {
                    console.error('[ImpulsePush] Failed to fetch VAPID key:', response.status, response.statusText);
                    throw new Error('Failed to fetch VAPID public key: ' + response.statusText);
                }
                const data = await response.json();
                this._vapidPublicKey = data.publicKey;
            }

            if (!this._vapidPublicKey) {
                console.error('[ImpulsePush] VAPID public key is empty.');
                return false;
            }

            const applicationServerKey = this.urlB64ToUint8Array(this._vapidPublicKey);

            // 2. Get or create browser push subscription
            let subscription = await swRegistration.pushManager.getSubscription();
            if (subscription) {
                // If subscription exists, verify applicationServerKey matches current server key
                if (subscription.options && subscription.options.applicationServerKey) {
                    const existingKeyArray = new Uint8Array(subscription.options.applicationServerKey);
                    let keyMatches = (existingKeyArray.length === applicationServerKey.length);
                    if (keyMatches) {
                        for (let i = 0; i < existingKeyArray.length; i++) {
                            if (existingKeyArray[i] !== applicationServerKey[i]) {
                                keyMatches = false;
                                break;
                            }
                        }
                    }
                    if (!keyMatches) {
                        console.log('[ImpulsePush] Server VAPID key updated. Refreshing push subscription token...');
                        await subscription.unsubscribe();
                        subscription = null;
                    }
                }
            }

            if (!subscription) {
                subscription = await swRegistration.pushManager.subscribe({
                    userVisibleOnly: true,
                    applicationServerKey: applicationServerKey
                });
            }

            const subJson = subscription.toJSON();
            const p256dh = (subJson.keys && subJson.keys.p256dh) || '';
            const auth = (subJson.keys && subJson.keys.auth) || '';

            // 3. Post subscription to backend (Upserts & associates with current logged in user)
            const payload = {
                endpoint: subscription.endpoint,
                keys: {
                    p256dh: p256dh,
                    auth: auth
                },
                userId: userId || '',
                userName: userName || ''
            };

            const saveResponse = await fetch(apiBase + 'api/push/subscribe', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (!saveResponse.ok) {
                console.error('[ImpulsePush] Failed to save subscription to server:', saveResponse.status);
                return false;
            }

            console.log('[ImpulsePush] Push subscription successfully registered for user:', userName);
            return true;
        } catch (error) {
            console.error('[ImpulsePush] Failed to sync push subscription:', error);
            return false;
        }
    },

    runDiagnostics: async function (userId, userName) {
        const steps = [];
        const log = (stepName, passed, details) => {
            steps.push({ step: stepName, passed: passed, details: details });
        };

        try {
            // Step 1: Secure Context
            const isSecure = window.isSecureContext;
            log('1. Secure Context', isSecure, isSecure ? 'window.isSecureContext is true' : 'Insecure origin (HTTPS or localhost required)');

            // Step 2: Browser API Support
            const hasSW = ('serviceWorker' in navigator);
            const hasNotif = ('Notification' in window);
            const hasPush = ('PushManager' in window);
            log('2. Browser APIs', hasSW && hasNotif, `ServiceWorker: ${hasSW}, Notification: ${hasNotif}, PushManager: ${hasPush}`);

            if (!hasSW || !hasNotif) {
                return { success: false, steps: steps };
            }

            // Step 3: Notification Permission
            let perm = Notification.permission;
            log('3. Initial Permission', perm === 'granted', `State: ${perm}`);

            if (perm !== 'granted') {
                try {
                    perm = await Notification.requestPermission();
                    log('3b. Requested Permission', perm === 'granted', `User selected: ${perm}`);
                    if (perm !== 'granted') return { success: false, steps: steps };
                } catch (e) {
                    log('3b. Request Permission Failed', false, e.name + ': ' + e.message);
                    return { success: false, steps: steps };
                }
            }

            // Step 4: Service Worker Registration
            const apiBase = this.getApiBaseUrl();
            const swUrl = apiBase + 'service-worker.js';
            let swRegistration = null;
            try {
                swRegistration = await navigator.serviceWorker.register(swUrl, { scope: apiBase });
                await navigator.serviceWorker.ready;
                log('4. Service Worker', true, `Active at: ${swUrl}`);
            } catch (e) {
                log('4. Service Worker Registration Failed', false, `${e.name}: ${e.message}`);
                return { success: false, steps: steps };
            }

            // Step 5: VAPID Key from Backend
            let vapidKey = this._vapidPublicKey;
            try {
                if (!vapidKey) {
                    const keyRes = await fetch(apiBase + 'api/push/vapid-public-key');
                    if (!keyRes.ok) throw new Error(`HTTP ${keyRes.status} ${keyRes.statusText}`);
                    const keyData = await keyRes.json();
                    vapidKey = keyData.publicKey;
                    this._vapidPublicKey = vapidKey;
                }
                log('5. VAPID Key Fetch', !!vapidKey, `Public Key received: ${vapidKey ? (vapidKey.substring(0, 16) + '...') : 'empty'}`);
            } catch (e) {
                log('5. VAPID Fetch Failed', false, `${e.name}: ${e.message}`);
                return { success: false, steps: steps };
            }

            // Step 6: PushManager Subscription with FCM
            let subscription = null;
            try {
                const appServerKey = this.urlB64ToUint8Array(vapidKey);
                subscription = await swRegistration.pushManager.getSubscription();
                if (!subscription) {
                    subscription = await swRegistration.pushManager.subscribe({
                        userVisibleOnly: true,
                        applicationServerKey: appServerKey
                    });
                }
                log('6. Push Subscription', !!subscription, `Token endpoint created (${subscription.endpoint.substring(0, 35)}...)`);
            } catch (e) {
                log('6. Push Subscription Failed', false, `${e.name}: ${e.message}`);
                return { success: false, steps: steps };
            }

            // Step 7: Post to /api/push/subscribe
            try {
                const subJson = subscription.toJSON();
                const payload = {
                    endpoint: subscription.endpoint,
                    keys: {
                        p256dh: (subJson.keys && subJson.keys.p256dh) || '',
                        auth: (subJson.keys && subJson.keys.auth) || ''
                    },
                    userId: userId || '',
                    userName: userName || ''
                };

                const postRes = await fetch(apiBase + 'api/push/subscribe', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!postRes.ok) {
                    const errText = await postRes.text();
                    log('7. SQL Database Save', false, `HTTP ${postRes.status}: ${errText}`);
                    return { success: false, steps: steps };
                }

                log('7. SQL Database Save', true, `Successfully registered in UserPushSubscriptions for ${userName || userId}`);
                return { success: true, steps: steps };
            } catch (e) {
                log('7. SQL Save Failed', false, `${e.name}: ${e.message}`);
                return { success: false, steps: steps };
            }
        } catch (globalEx) {
            log('Diagnostics Unexpected Error', false, globalEx.message || globalEx);
            return { success: false, steps: steps };
        }
    }
};
