// sw.js (Adom's Lab 統合版 v10: 画像名修正 & オフラインアイコン対応)
const CACHE_NAME = 'adom-lab-v10'; // ★バージョンを10に更新
const INITIAL_ASSETS = [
    // --- アプリ本体 (HTML) ---
    './gotoshitrash.html',
    './teachers.html',
    './gotostarview.html',
    './mission.html',

    // --- アイコン画像 (ホーム画面用) ---
    // ★ファイル名を修正しました
    './gotobaramon.png',   // ゴミ分別
    './plannavi.png',      // ルート作成 (Mission Ctrl)
    './gotostarview.png',  // 星空アプリ

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

    // ▼▼▼ 【重要】アイコンの「実体ファイル」を追加 ▼▼▼
    // これがないと、オフライン時にボタンが「□」になったり消えたりします
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.woff2',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-brands-400.woff2', 
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-regular-400.woff2',

    // --- 機能特化ライブラリ ---
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js',

    // --- フォント (Google Fonts) ---
    'https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap',
    'https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap'
];

// 1. インストール
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            // エラーが出ても止まらないように、1つずつ追加を試みる
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
