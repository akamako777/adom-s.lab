// sw.js (Adom's Lab 統合版 v13: クラス管理ツールPro対応版)
// ローカルファイル(JS)とCDN(CSS/Webfonts)、どちらを使っていてもキャッシュするように設定しています

const CACHE_NAME = 'adom-lab-v13-teachers-pro'; // ★更新のためバージョン名を変更(v12 -> v13)
const INITIAL_ASSETS = [
    // ---------------------------
    // 1. アプリ本体 (HTML)
    // ---------------------------
    './gotoshitrash.html',
    './teachers.html',
    './gotostarview.html',
    './mission.html',

    // ---------------------------
    // 2. マニフェストとアイコン
    // ---------------------------
    './manifest_trash.json',
    './gotobaramon.png',
    './plannavi.png',
    './gotostarview.png',
    './manifest_teachers.json', // ★追加：クラス管理ツール用マニフェスト
    './teachers.png',           // ★追加：クラス管理ツール用アイコン

    // ---------------------------
    // 3. ローカル用ライブラリ
    // (gotoshitrash.html や teachers.html が使うファイル)
    // ---------------------------
    './react.production.min.js',
    './react-dom.production.min.js',
    './babel.min.js',
    './tailwindcss.js',
    './fontawesome.min.js',
    './xlsx.full.min.js',       // ★追加：Excel出力用 (クラス管理ツール用)
    './jszip.min.js',           // ★追加：ZIP圧縮用 (クラス管理ツールの写真一括ダウンロード用)

    // ---------------------------
    // 4. 【旧】CDN用ライブラリ
    // (他のアプリがまだこれを使っている場合の保険)
    // ---------------------------
    // React & ReactDOM (cdnjs v18.2.0)
    'https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js',
    // Babel (cdnjs v7.23.5)
    'https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.5/babel.min.js',
    
    // React & ReactDOM (unpkg v18 - teachers.htmlなどが使用)
    'https://unpkg.com/react@18/umd/react.production.min.js',
    'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js',
    // Babel (unpkg)
    'https://unpkg.com/@babel/standalone/babel.min.js',
    
    // TailwindCSS (CDN版)
    'https://cdn.tailwindcss.com',

    // ---------------------------
    // 5. 共通ツール・フォント・機能ライブラリ
    // (他のアプリのためにCDN版FontAwesomeも残しておきます)
    // ---------------------------
    // FontAwesome CSS (CDN)
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
    
    // FontAwesome Webfonts (CDN版を使っているアプリ用)
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.woff2',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-brands-400.woff2', 
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-regular-400.woff2',

    // Excel処理 (xlsx) - ※CDN版も他のアプリの保険として残しています
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    // QRコード (qrious)
    'https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js',

    // Google Fonts (日本語フォントなど)
    'https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap',
    'https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap',
    // Trashアプリで使っているNoto Sans JP
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
                        // ファイルが無い、またはネット切れ等の場合もエラーで止めずに無視する
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
    // http/https 以外のリクエストは無視
    if (!event.request.url.startsWith('http')) return;

    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            // 1. キャッシュにあればそれを返す (ローカル・CDN問わず)
            if (cachedResponse) {
                return cachedResponse;
            }

            // 2. なければネットに取りに行く
            return fetch(event.request).then((networkResponse) => {
                // 正常なレスポンスでなければそのまま返す
                if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic' && networkResponse.type !== 'cors') {
                    return networkResponse;
                }

                // 3. ネットから取れたら、次回のためにキャッシュに保存しておく
                const responseToCache = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseToCache);
                });

                return networkResponse;
            }).catch(() => {
                // 4. オフラインで、キャッシュも無い場合
                return null; // またはオフライン用画像を返すなど
            });
        })
    );
});
