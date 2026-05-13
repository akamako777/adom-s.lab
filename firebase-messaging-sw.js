// Firebaseの裏方専用ライブラリを読み込み
importScripts('https://www.gstatic.com/firebasejs/10.8.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.8.1/firebase-messaging-compat.js');

// ★取得した1つ目のカギ（設定コード）
firebase.initializeApp({
  apiKey: "AIzaSyDDi3IdcvyBh4R3nh7_bMQkabpWfcZpPFk",
  authDomain: "goto-gomi-app.firebaseapp.com",
  projectId: "goto-gomi-app",
  storageBucket: "goto-gomi-app.firebasestorage.app",
  messagingSenderId: "730896181153",
  appId: "1:730896181153:web:eeaa805011a4a807250de0"
});

// バックグラウンドで通知を受け取る準備
const messaging = firebase.messaging();