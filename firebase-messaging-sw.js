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
const ICON_URL = 'https://ghwkdwjd-debug.github.io/logipan-report/icon.png';
const APP_URL = 'https://ghwkdwjd-debug.github.io/logipan-report/';

// [중요] 중복 방지를 위해 onBackgroundMessage만 사용 (push 이벤트는 SDK가 알아서 처리)
// onBackgroundMessage 안에서 알림을 직접 띄우므로 push 이벤트 핸들러는 제거함
messaging.onBackgroundMessage(payload => {
    console.log('[FCM-SW] 백그라운드 메시지:', JSON.stringify(payload));

    // data 블록에서 정보 추출
    const title = payload.data?.title || payload.notification?.title || '로지판';
    const body = payload.data?.body || payload.notification?.body || '';

    return self.registration.showNotification(title, {
        body: body,
        icon: ICON_URL,
        badge: ICON_URL,
        tag: 'logipan-' + Date.now(),
        requireInteraction: false,
        renotify: true,
        data: payload.data || {}
    });
});

// 알림 클릭 시 → 사이트 열기
self.addEventListener('notificationclick', event => {
    event.notification.close();
    event.waitUntil(
        clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clientList => {
            for (const client of clientList) {
                if (client.url.includes('logipan-report') && 'focus' in client) {
                    return client.focus();
                }
            }
            if (clients.openWindow) {
                return clients.openWindow(APP_URL);
            }
        })
    );
});

// 서비스워커 즉시 활성화
self.addEventListener('install', event => {
    self.skipWaiting();
});

self.addEventListener('activate', event => {
    event.waitUntil(clients.claim());
});
