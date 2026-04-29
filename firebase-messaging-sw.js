// 로지판 - Firebase Cloud Messaging 서비스 워커
// 앱 닫혀있어도 알림이 폰에 오게 하는 핵심 파일

importScripts('https://www.gstatic.com/firebasejs/8.10.1/firebase-app.js');
importScripts('https://www.gstatic.com/firebasejs/8.10.1/firebase-messaging.js');

firebase.initializeApp({
    apiKey: "AIzaSyBYJfQD7Jkd9Jecyu27Owy8yGPrwg3tg80",
    authDomain: "logipan-2026.firebaseapp.com",
    projectId: "logipan-2026",
    storageBucket: "logipan-2026.firebasestorage.app",
    messagingSenderId: "650344159406",
    appId: "1:650344159406:web:ff736627926763dae6353f"
});

const messaging = firebase.messaging();

// 백그라운드 메시지 수신 (앱 꺼져있을 때)
messaging.onBackgroundMessage(payload => {
    console.log('[FCM-SW] 백그라운드 메시지 수신:', payload);
    const notificationTitle = payload.notification?.title || '로지판';
    const notificationOptions = {
        body: payload.notification?.body || '새 알림이 있습니다',
        icon: '/logipan-report/icon.png',
        badge: '/logipan-report/icon.png',
        tag: 'logipan-notification',
        requireInteraction: false,
        data: payload.data || {}
    };
    self.registration.showNotification(notificationTitle, notificationOptions);
});

// 알림 클릭 시 → 사이트 열기
self.addEventListener('notificationclick', event => {
    event.notification.close();
    event.waitUntil(
        clients.matchAll({ type: 'window' }).then(clientList => {
            for (const client of clientList) {
                if (client.url.includes('logipan-report') && 'focus' in client) {
                    return client.focus();
                }
            }
            if (clients.openWindow) {
                return clients.openWindow('https://ghwkdwjd-debug.github.io/logipan-report/');
            }
        })
    );
});
