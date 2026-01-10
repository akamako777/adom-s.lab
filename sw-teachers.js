// sw-teachers.js (クラス編成アプリ専用の倉庫番)
const CACHE_NAME = 'teachers-manager-pro-v1';
const ASSETS = [
    './teachers.html', // アプリ本体
    'https://cdn.tailwindcss.com', // デザイン
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css', // アイコン
    // もしReactを使っているなら以下もキャッシュされます
    'https://unpkg.com/react@18/umd/react.production.min.js',
    'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js',
    'https://unpkg.com/@babel/standalone/babel.min.js'
];

// インストール時にデータをキャッシュ（保存）
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
