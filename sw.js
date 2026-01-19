// sw.js (Adom's Lab 統合版 v7: 星空アプリ "Highest Masterpiece" 追加)
const CACHE_NAME = 'adom-lab-v7'; // ★バージョンを v7 に更新！
const INITIAL_ASSETS = [
    // --- アプリ本体 ---
    './gotoshitrash.html', // ゴミ分別アプリ
    './teachers.html',     // 先生ツール
    './gotostarview.html', // ★今回追加した最高傑作（星空アプリ）

    // --- 必須エンジン (React/Babel) ※これがないとオフラインで動きません ---
    'https://unpkg.com/react@18/umd/react.production.min.js',
    'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js',
    'https://unpkg.com/@babel/standalone/babel.min.js',

    // --- 共通ツール・デザイン (Tailwind, FontAwesome) ---
    'https://cdn.tailwindcss.com',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',

    // --- 星空アプリ専用フォント・ツール ---
    'https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;700;900&display=swap', // 丸ゴシック
    'https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap', // デジタル時計風フォント
    'https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js'      // QRコード生成
];

// 1. インストール（アプリ本体と基本ツールの確保）
self.addEventListener('install', (event) => {
    self.skipWaiting(); // 待機せずにすぐ新しいバージョンを適用
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(INITIAL_ASSETS);
        })
    );
});

// 2. 有効化（古いバージョンのキャッシュをお掃除）
self.addEventListener('activate', (event) => {
    event.waitUntil(clients.claim()); // すぐにコントロール開始
    event.waitUntil(
        caches.keys().then((keyList) => {
            return Promise.all(keyList.map((key) => {
                // 新しいバージョン(v7)以外は削除する
                if (key !== CACHE_NAME) {
                    return caches.delete(key);
                }
            }));
        })
    );
});

// 3. 通信の横取り & 自動保存（オフライン対応の要）
self.addEventListener('fetch', (event) => {
    // http/https 以外の通信（chrome-extension等）は無視
    if (!event.request.url.startsWith('http')) return;

    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            // A. スマホの中に保存データがあれば、それを返す（オフライン成功！）
            if (cachedResponse) {
                return cachedResponse;
            }

            // B. なければインターネットに取りに行く
            return fetch(event.request).then((networkResponse) => {
                // エラーや無効なデータならそのまま返す
                if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic' && networkResponse.type !== 'cors') {
                    return networkResponse;
                }

                // C. ネットから取れたデータは、次回のためにコピーして保存しておく（自動学習）
                const responseToCache = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseToCache);
                });

                return networkResponse;
            }).catch(() => {
                // ネットも繋がらず、キャッシュもない場合（エラー回避）
                return null;
            });
        })
    );
});
