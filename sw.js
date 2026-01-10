// sw.js (Service Worker)
const CACHE_NAME = 'goto-gomi-v1';
const ASSETS = [
    './gotoshitrash.html', // アプリ本体
    'https://cdn.tailwindcss.com', // デザインツール
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css', // アイコン
    'https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap' // フォント
];

// インストール時にデータをキャッシュする
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(ASSETS);
        })
    );
});

// オフライン時にキャッシュからデータを出す
self.addEventListener('fetch', (event) => {
    event.respondWith(
        caches.match(event.request).then((response) => {
            return response || fetch(event.request);
        })
    );
});
