// sw.js (強力版: 自動保存機能付き)
const CACHE_NAME = 'goto-gomi-v3'; // バージョンを更新
const INITIAL_ASSETS = [
    './gotoshitrash.html' // アプリ本体だけは絶対に確保
];

// 1. インストール（本体の確保）
self.addEventListener('install', (event) => {
    self.skipWaiting(); // 待機せずに即更新
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(INITIAL_ASSETS);
        })
    );
});

// 2. 有効化（古いデータの削除）
self.addEventListener('activate', (event) => {
    event.waitUntil(clients.claim()); // すぐに制御開始
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

// 3. 通信の横取り（ここが重要！）
self.addEventListener('fetch', (event) => {
    // http以外のリクエスト（chrome-extension等）は無視
    if (!event.request.url.startsWith('http')) return;

    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            // A. キャッシュにあれば、それを返す（オフライン対応）
            if (cachedResponse) {
                return cachedResponse;
            }

            // B. なければインターネットに取りに行く
            return fetch(event.request).then((networkResponse) => {
                // 有効なレスポンスでなければそのまま返す
                if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic' && networkResponse.type !== 'cors') {
                    return networkResponse;
                }

                // C. 取得したデータを、次回のために「コピーして保存」しておく
                const responseToCache = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseToCache);
                });

                return networkResponse;
            }).catch(() => {
                // オフラインで、キャッシュにもない場合（エラー回避）
                return null;
            });
        })
    );
});
