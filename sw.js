// sw.js (Adom's Lab 統合版: 自動保存機能付き)
const CACHE_NAME = 'adom-lab-v4'; // バージョン更新
const INITIAL_ASSETS = [
    './gotoshitrash.html', // ゴミ分別アプリ
    './teachers.html'      // 先生ツール（ここを追加！）
];

// 1. インストール（アプリ本体の確保）
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(INITIAL_ASSETS);
        })
    );
});

// 2. 有効化（古いデータの掃除）
self.addEventListener('activate', (event) => {
    event.waitUntil(clients.claim());
    event.waitUntil(
        caches.keys().then((keyList) => {
            return Promise.all(keyList.map((key) => {
                if (key !== CACHE_NAME) {
                    return caches.delete(key);
                }
            }));
        })
    );
});

// 3. 通信の横取り & 自動保存（ここが最強）
self.addEventListener('fetch', (event) => {
    // http/https 以外は無視
    if (!event.request.url.startsWith('http')) return;

    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            // A. キャッシュにあれば、それを返す（オフライン対応）
            if (cachedResponse) {
                return cachedResponse;
            }

            // B. なければインターネットに取りに行く
            return fetch(event.request).then((networkResponse) => {
                // 有効なデータじゃなければそのまま返す
                if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic' && networkResponse.type !== 'cors') {
                    return networkResponse;
                }

                // C. 取得した「重たいライブラリ（React, Babel等）」も自動保存！
                const responseToCache = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseToCache);
                });

                return networkResponse;
            }).catch(() => {
                return null; // エラー回避
            });
        })
    );
});
