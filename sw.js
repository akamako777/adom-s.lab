// sw.js (Adom's Lab 統合版 v9: Mission Ctrl 追加対応)
const CACHE_NAME = 'adom-lab-v9'; // ★バージョン更新！
const INITIAL_ASSETS = [
    // --- アプリ本体 (HTML) ---
    './gotoshitrash.html', // ゴミ分別
    './teachers.html',     // 先生ツール
    './gotostarview.html', // 星空アプリ
    './mission.html',      // ★追加：ルート作成アプリ

    // --- アイコン画像 (オフライン表示用) ---
    // ※ファイル名が確定しているものだけ記述。
    // ※もしファイル名が違う場合は修正してください。
    './gotobaramon.png',   // ゴミ分別アイコン
    './plannavi.png', // ルート作成アイコン
    './gotostarview.png',    // 星空アイコン

    // --- 必須エンジン (React/Babel - cdnjs版) ---
    'https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.5/babel.min.js',

    // --- 必須エンジン (React/Babel - unpkg版) ---
    'https://unpkg.com/react@18/umd/react.production.min.js',
    'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js',
    'https://unpkg.com/@babel/standalone/babel.min.js',

    // --- 共通ツール・デザイン ---
    'https://cdn.tailwindcss.com',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',

    // --- 機能特化ライブラリ ---
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', // Excel処理
    'https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js',    // QRコード

    // --- フォント ---
    'https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap',
    'https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap'
];

// 1. インストール
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            // エラーが出ても止まらないように、1つずつ追加を試みる
            // (アイコン画像が無い場合などにインストール失敗するのを防ぐため)
            return Promise.all(
                INITIAL_ASSETS.map(url => {
                    return cache.add(url).catch(err => {
                        console.log('Cache failed for:', url);
                    });
                })
            );
        })
    );
});

// 2. 有効化（旧バージョンの削除）
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

// 3. 通信の横取り & 自動保存
self.addEventListener('fetch', (event) => {
    if (!event.request.url.startsWith('http')) return;

    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            if (cachedResponse) {
                return cachedResponse;
            }
            return fetch(event.request).then((networkResponse) => {
                if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic' && networkResponse.type !== 'cors') {
                    return networkResponse;
                }
                const responseToCache = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseToCache);
                });
                return networkResponse;
            }).catch(() => {
                return null;
            });
        })
    );
});
