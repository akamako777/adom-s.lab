// sw.js (Adom's Lab 統合版 v14: パワハラ撲滅アプリ対応版)
// ローカルファイル(JS)とCDN(CSS/Webfonts)、どちらを使っていてもキャッシュするように設定しています

const CACHE_NAME = 'adom-lab-v14-health-def'; // ★更新のためバージョン名を変更(v13 -> v14)

const INITIAL_ASSETS = [
    // ---------------------------
    // 1. アプリ本体 (HTML)
    // ---------------------------
    './gotoshitrash.html',
    './teachers.html',
    './gotostarview.html',
    './mission.html',
    './st.healthdefrec.html', // ★追加：パワハラ撲滅アプリ本体

    // ---------------------------
    // 2. マニフェストとアイコン (PWA用)
    // ---------------------------
    // ※st.healthdefrec.htmlはステルス性を高めるためマニフェストは登録しません
    './manifest_trash.json',
    './gotobaramon.png',
    './plannavi.png',
    './gotostarview.png',
    './manifest_teachers.json', 
    './teachers.png',           

    // ---------------------------
    // 3. ローカル用ライブラリ
    // ---------------------------
    './react.production.min.js',
    './react-dom.production.min.js',
    './babel.min.js',
    './tailwindcss.js',
    './fontawesome.min.js',
    './xlsx.full.min.js',       
    './jszip.min.js',           

    // ---------------------------
    // 4. 【旧】CDN用ライブラリ
    // ---------------------------
    'https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.5/babel.min.js',
    
    'https://unpkg.com/react@18/umd/react.production.min.js',
    'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js',
    'https://unpkg.com/@babel/standalone/babel.min.js',
    
    'https://cdn.tailwindcss.com',

    // ---------------------------
    // 5. 共通ツール・フォント・機能ライブラリ
    // ---------------------------
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
    
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.woff2',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-brands-400.woff2', 
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-regular-400.woff2',

    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js',

    'https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap',
    'https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap',
    'https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap'
];

// --- インストール処理 (登録時に全ファイルをキャッシュへ) ---
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            // 1つ失敗しても他は保存する「安全策」で登録
            return Promise.all(
                INITIAL_ASSETS.map(url => {
                    return cache.add(url).catch(err => {
                        console.log('Skipped:', url); 
                    });
                })
            );
        })
    );
});

// --- 有効化処理 (古いバージョンのキャッシュ削除) ---
self.addEventListener('activate', (event) => {
    event.waitUntil(clients.claim());
    event.waitUntil(
        caches.keys().then((keyList) => {
            return Promise.all(keyList.map((key) => {
                // バージョン名が違う古いキャッシュは消す
                if (key !== CACHE_NAME) {
                    return caches.delete(key);
                }
            }));
        })
    );
});

// --- 通信処理 (キャッシュ優先) ---
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
