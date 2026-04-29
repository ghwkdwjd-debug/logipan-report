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

// [수정] 백그라운드 메시지 수신 + data 메시지도 처리
messaging.onBackgroundMessage(payload => {
    console.log('[FCM-SW] 백그라운드 메시지 수신:', JSON.stringify(payload));

    // notification 또는 data 둘 중 하나에서 정보 추출
    const title = payload.notification?.title || payload.data?.title || '로지판';
    const body = payload.notification?.body || payload.data?.body || '새 알림이 있습니다';

    const notificationOptions = {
        body: body,
        icon: ICON_URL,
        badge: ICON_URL,
        tag: 'logipan-' + Date.now(),
        requireInteraction: false,
        renotify: true,
        data: payload.data || {}
    };
    return self.registration.showNotification(title, notificationOptions);
});

// [추가] iOS PWA 보강 - push 이벤트도 직접 처리
// onBackgroundMessage가 어떤 이유로 안 트리거될 때를 위한 안전망
self.addEventListener('push', event => {
    console.log('[FCM-SW] push 이벤트 발생');
    if (!event.data) {
        console.log('[FCM-SW] push 이벤트에 data 없음');
        return;
    }

    let data;
    try {
        data = event.data.json();
        console.log('[FCM-SW] push 데이터:', JSON.stringify(data));
    } catch (e) {
        const text = event.data.text();
        console.log('[FCM-SW] push 텍스트:', text);
        data = { notification: { title: '로지판', body: text } };
    }

    // FCM이 보내는 형식: { notification: {...}, data: {...} } 또는 { data: {...} }
    const title = data.notification?.title || data.data?.title || data.title || '로지판';
    const body = data.notification?.body || data.data?.body || data.body || '새 알림';

    event.waitUntil(
        self.registration.showNotification(title, {
            body: body,
            icon: ICON_URL,
            badge: ICON_URL,
            tag: 'logipan-push-' + Date.now(),
            requireInteraction: false,
            renotify: true,
        })
    );
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
    console.log('[FCM-SW] install');
    self.skipWaiting();
});

self.addEventListener('activate', event => {
    console.log('[FCM-SW] activate');
    event.waitUntil(clients.claim());
});
