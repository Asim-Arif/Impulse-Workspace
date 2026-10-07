// Impulse ERP Web Push Service Worker
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});

self.addEventListener('push', (event) => {
    if (!event.data) {
        return;
    }

    let payload = {};
    try {
        payload = event.data.json();
    } catch (e) {
        payload = {
            title: 'Impulse ERP',
            body: event.data.text(),
            data: { url: '/' }
        };
    }

    const title = payload.title || 'Impulse ERP Notification';
    const body = payload.body || payload.message || '';
    const icon = payload.icon || '/favicon.ico';
    const badge = payload.badge || '/favicon.ico';
    const url = (payload.data && payload.data.url) || payload.actionUrl || '/';

    const options = {
        body: body,
        icon: icon,
        badge: badge,
        vibrate: [100, 50, 100],
        data: {
            url: url
        },
        tag: 'impulse-notification-' + (payload.id || Date.now()),
        renotify: true,
        requireInteraction: false
    };

    event.waitUntil(
        self.registration.showNotification(title, options)
    );
});

self.addEventListener('notificationclick', (event) => {
    event.notification.close();

    const targetUrl = (event.notification.data && event.notification.data.url) ? event.notification.data.url : '/';

    event.waitUntil(
        clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
            // If an active app window is already open, focus it and navigate
            for (let i = 0; i < clientList.length; i++) {
                const client = clientList[i];
                if ('focus' in client) {
                    client.focus();
                    if ('navigate' in client && targetUrl !== '/') {
                        client.navigate(targetUrl);
                    }
                    return;
                }
            }
            // Otherwise open a new window
            if (clients.openWindow) {
                return clients.openWindow(targetUrl);
            }
        })
    );
});
