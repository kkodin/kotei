// 工事日報：アプリとしてインストールできるようにするだけ（控えは持たない。いつも最新を読む）
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
