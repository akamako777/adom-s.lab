// sw.js (Adom's Lab 統合版 v16: ハイブリッドPWA対応版)

const CACHE_NAME = 'adom-lab-v16-hybrid'; // ★v16に更新（古いバグキャッシュを強制削除）

const INITIAL_ASSETS = [
    // ---------------------------
    // 1. アプリ本体 (HTML)
    // ---------------------------
    './gotoshitrash.html',
    './teachers.html',
    './gotostarview.html',
    './mission.html',
    './st.healthdefrec.html',
    './forocrpdf.html',
    './englishprint.html', // 英語プリント
    './englishwrite.html', // 英語プリント（新）
    './oldphotoscan.html', // レトロスキャン

    // ---------------------------
    // 2. マニフェストとアイコン (PWA用)
    // ---------------------------
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
    // 4. CDN用ライブラリ ＆ 新規追加ライブラリ
    // ---------------------------
    'https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.5/babel.min.js',
    
    'https://unpkg.com/react@18/umd/react.production.min.js',
    'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js',
    'https://unpkg.com/@babel/standalone/babel.min.js',
    
    'https://cdn.tailwindcss.com',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', 

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

// --- インストール処理 ---
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return Promise.all(
                INITIAL_ASSETS.map(url => {
                    return cache.add(url).catch(err => {
                        console.log('Skipped caching:', url); 
                    });
                })
            );
        })
    );
});

// --- 有効化処理 (古いキャッシュの削除) ---
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

// --- 通信処理 (ハイブリッドルーティング) ---
self.addEventListener('fetch', (event) => {
    if (!event.request.url.startsWith('http')) return;

    // ★重要変更：HTMLファイルへのアクセスは「ネットワーク優先 (Network First)」にする
    // （常に最新のアプリ画面を取得し、Aを開いてBが出るバグを防止）
    if (event.request.mode === 'navigate' || (event.request.headers.get('accept') && event.request.headers.get('accept').includes('text/html'))) {
        event.respondWith(
            fetch(event.request).then((networkResponse) => {
                // 通信成功：最新HTMLをキャッシュに上書きして返す
                return caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, networkResponse.clone());
                    return networkResponse;
                });
            }).catch(() => {
                // オフライン時：保存してあるキャッシュを返す
                return caches.match(event.request);
            })
        );
        return;
    }

    // ★それ以外のファイル（画像、CSS、JS）は「キャッシュ優先 (Cache First)」で爆速読み込み
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
            }).catch(() => null);
        })
    );
});
